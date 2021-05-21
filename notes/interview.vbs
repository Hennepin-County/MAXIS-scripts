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


const ref_number					= 0
const access_denied					= 1
const full_name_const				= 2
const last_name_const				= 3
const first_name_const				= 4
const mid_initial					= 5
const age							= 6
const date_of_birth					= 7
const ssn							= 8
const ssn_verif						= 9
const birthdate_verif				= 10
const gender						= 11
const race							= 12
const spoken_lang					= 13
const written_lang					= 14
const interpreter					= 15
const alias_yn						= 16
const ethnicity_yn					= 17
const id_verif						= 18
const rel_to_applcnt				= 19
const cash_minor					= 20
const snap_minor					= 21
const marital_status				= 22
const spouse_ref					= 23
const spouse_name					= 24
const last_grade_completed 			= 25
const citizen						= 26
const other_st_FS_end_date 			= 27
const in_mn_12_mo					= 28
const residence_verif				= 29
const mn_entry_date					= 30
const former_state					= 31
const fs_pwe						= 32
const button_one					= 33
const button_two					= 34
const clt_has_sponsor				= 35
const client_verification			= 36
const client_verification_details	= 37
const client_notes					= 38
const intend_to_reside_in_mn		= 39
const race_a_checkbox				= 40
const race_b_checkbox				= 41
const race_n_checkbox				= 42
const race_p_checkbox				= 43
const race_w_checkbox				= 44
const snap_req_checkbox				= 45
const cash_req_checkbox				= 46
const emer_req_checkbox				= 47
const none_req_checkbox				= 48
const ssn_no_space					= 49
const edrs_msg						= 50
const edrs_match					= 51
const edrs_notes 					= 52
const last_const					= 53

Dim HH_MEMB_ARRAY()
ReDim HH_MEMB_ARRAY(last_const, 0)



'===========================================================================================================================

'FUNCTIONS =================================================================================================================

'
' class mx_hh_member
'
' 	public access_denied
' 	public selected
' 	'stuff about the members
' 	public first_name
' 	public last_name
' 	public mid_initial
' 	public other_names
' 	public date_of_birth
' 	public age
' 	public ref_number
' 	public ssn
' 	public ssn_verif
' 	public birthdate_verif
' 	public gender
' 	public id_verif
' 	public rel_to_applcnt
' 	public race
' 	public race_a_checkbox
' 	public race_b_checkbox
' 	public race_n_checkbox
' 	public race_p_checkbox
' 	public race_w_checkbox
' 	public snap_minor
' 	public cash_minor
' 	public written_lang
' 	public spoken_lang
' 	public interpreter
' 	public alias_yn
' 	public ethnicity_yn
'
' 	public marital_status
' 	public spouse_ref
' 	public spouse_name
' 	public last_grade_completed
' 	public citizen
' 	public other_st_FS_end_date
' 	public in_mn_12_mo
' 	public residence_verif
' 	public mn_entry_date
' 	public former_state
' 	public intend_to_reside_in_mn
'
' 	public parents_in_home
' 	public parents_in_home_notes
' 	public parent_one_name
' 	public parent_one_type
' 	public parent_one_verif
' 	public parent_one_in_home
' 	public parent_two_name
' 	public parent_two_type
' 	public parent_two_verif
' 	public parent_two_in_home
'
' 	public pare_exists
' 	public pare_child_ref_nbr
' 	public pare_child_name
' 	public pare_child_member_index
' 	public pare_relationship_type
' 	public pare_verification
'
' 	public remo_exists
' 	public left_hh_date
' 	public left_hh_reason
' 	public left_hh_expected_return_date
' 	public left_hh_expected_return_verif
' 	public left_hh_actual_return_date
' 	public left_hh_HC_temp_out_of_state
' 	public left_hh_date_reported
' 	public left_hh_12_months_or_more
'
' 	public adme_exists
' 	public adme_arrival_date
' 	public adme_cash_date
' 	public adme_emer_date
' 	public adme_snap_date
' 	public adme_within_12_months
'
' 	public imig_exists
' 	public imig_status
' 	public us_entry_date
' 	public imig_status_date
' 	public imig_status_verif
' 	public lpr_adj_from
' 	public nationality
' 	public nationality_detail
' 	public alien_id_nbr
' 	public imig_active_doc
' 	public imig_recvd_doc
' 	public imig_q_2_required
' 	public imig_q_4_required
' 	public imig_q_5_required
' 	public imig_clt_current_doc
' 	public imig_doc_on_file
' 	public imig_save_completed
' 	public imig_prev_status
'
' 	public new_imig_status
' 	public new_us_entry_date
' 	public new_imig_status_date
' 	public new_imig_status_verif
' 	public new_lpr_adj_from
' 	public new_nationality
' 	public new_nationality_detail
' 	public new_imig_active_doc
' 	public new_imig_recvd_doc
' 	public new_imig_clt_current_doc
' 	public new_imig_doc_on_file
' 	public new_imig_save_completed
' 	public new_imig_prev_status
' 	public new_spon_name
' 	public new_spon_street
' 	public new_spon_city
' 	public new_spon_state
' 	public new_spon_zip
' 	public new_spon_phone
' 	public new_spon_gross_income
' 	public new_spon_income_freq
' 	public new_spon_spouse_name
' 	public new_spon_spouse_income
' 	public new_spon_spouse_income_freq
'
' 	' public ans_us_entry_date
' 	' public ans_nationality
' 	' public ans_nationality_detail
' 	' public ans_imig_status
' 	' public ans_imig_prev_status
' 	' public ans_imig_status_date
' 	' public ans_imig_clt_current_doc
' 	' public ans_imig_doc_on_file
' 	' public ans_imig_save_completed
' 	' public ans_clt_has_sponsor
' 	' public ans_spon_name
' 	' public ans_live_with_spon
' 	' public ans_spon_street
' 	' public ans_spon_city
' 	' public ans_spon_state
' 	' public ans_spon_zip
' 	' public ans_spon_phone
' 	' public ans_spon_gross_income
' 	' public ans_spon_income_freq
' 	' public ans_spon_married_yn
' 	' public ans_spon_children_yn
' 	' public ans_spon_spouse_name
' 	' public ans_spon_spouse_income
' 	' public ans_spon_spouse_income_freq
' 	' public ans_spon_numb_children
' 	' public ans_spon_hh_notes
'
' 	public spon_exists
' 	public clt_has_sponsor
' 	' public ask_about_spon
' 	public spon_type
' 	public spon_verif
' 	public spon_name
' 	public spon_street
' 	public spon_city
' 	public spon_state
' 	public spon_zip
' 	public spon_phone
' 	public spon_cash_retro_income
' 	public spon_cash_prosp_income
' 	public spon_cash_assets
' 	public spon_snap_retro_income
' 	public spon_snap_prosp_income
' 	public spon_snap_assets
' 	public spon_spouse
' 	public spon_hh_size
' 	public spon_numb_children
' 	public spon_possible_indigent_exemption
' 	public spon_gross_income
' 	public spon_spouse_income
' 	public live_with_spon
' 	public spon_income_freq
' 	public spon_spouse_income_freq
' 	public spon_married_yn
' 	public spon_children_yn
' 	public spon_hh_notes
' 	public spon_spouse_name
'
' 	public disa_exists
' 	public disa_begin_date
' 	public disa_end_date
' 	public disa_cert_begin_date
' 	public disa_cert_end_date
' 	public cash_disa_status
' 	public cash_disa_verif
' 	public fs_disa_status
' 	public fs_disa_verif
' 	public hc_disa_status
' 	public hc_disa_verif
' 	public disa_waiver
' 	public disa_1619
' 	public disa_detail
' 	public mof_file
' 	public mof_detail
' 	public mof_end_date
' 	public iaa_file
' 	public iaa_received_date
' 	public iaa_complete
' 	public disa_review
'
' 	public fs_pwe
' 	public wreg_exists
'
' 	public schl_exists
' 	public school_status
' 	public school_grade
' 	public school_name
' 	public school_verif
' 	public school_type
' 	public school_district
' 	public kinder_start_date
' 	public grad_date
' 	public grad_date_verif
' 	public school_funding
' 	public school_elig_status
' 	public higher_ed
'
' 	public stin_exists
' 	public total_stin
' 	public stin_type_array
' 	public stin_amount_array
' 	public stin_avail_date_array
' 	public stin_months_cov_array
' 	public stin_verif_array
'
' 	public stec_exists
' 	public total_stec
' 	public stec_type_array
' 	public stec_amount_array
' 	public stec_months_cov_array
' 	public stec_verif_array
' 	public stec_earmarked_amount_array
' 	public stec_earmarked_months_cov_array
'
' 	public shel_exists
' 	public shel_summary
' 	public shel_hud_subsidy_yn
' 	public shel_shared_yn
' 	public shel_paid_to
' 	public shel_retro_rent_amount
' 	public shel_retro_rent_verif
' 	public shel_retro_lot_rent_amount
' 	public shel_retro_lot_rent_verif
' 	public shel_retro_mortgage_amount
' 	public shel_retro_mortgage_verif
' 	public shel_retro_insurance_amount
' 	public shel_retro_insurance_verif
' 	public shel_retro_taxes_amount
' 	public shel_retro_taxes_verif
' 	public shel_retro_room_amount
' 	public shel_retro_room_verif
' 	public shel_retro_garage_amount
' 	public shel_retro_garage_verif
' 	public shel_retro_subsidy_amount
' 	public shel_retro_subsidy_verif
'
' 	public shel_prosp_rent_amount
' 	public shel_prosp_rent_verif
' 	public shel_prosp_lot_rent_amount
' 	public shel_prosp_lot_rent_verif
' 	public shel_prosp_mortgage_amount
' 	public shel_prosp_mortgage_verif
' 	public shel_prosp_insurance_amount
' 	public shel_prosp_insurance_verif
' 	public shel_prosp_taxes_amount
' 	public shel_prosp_taxes_verif
' 	public shel_prosp_room_amount
' 	public shel_prosp_room_verif
' 	public shel_prosp_garage_amount
' 	public shel_prosp_garage_verif
' 	public shel_prosp_subsidy_amount
' 	public shel_prosp_subsidy_verif
'
' 	public coex_exists
' 	public coex_support_verif
' 	public coex_support_retro_amount
' 	public coex_support_prosp_amount
' 	public coex_support_hc_est_amount
' 	public coex_alimony_verif
' 	public coex_alimony_retro_amount
' 	public coex_alimony_prosp_amount
' 	public coex_alimony_hc_est_amount
' 	public coex_tax_dep_verif
' 	public coex_tax_dep_retro_amount
' 	public coex_tax_dep_prosp_amount
' 	public coex_tax_dep_hc_est_amount
' 	public coex_other_verif
' 	public coex_other_retro_amount
' 	public coex_other_prosp_amount
' 	public coex_other_hc_est_amount
' 	public coex_total_retro_amount
' 	public coex_total_prosp_amount
' 	public coex_total_hc_est_amount
' 	public coex_change_in_financial_circumstances
'
' 	public stwk_exists
' 	public stwk_employer
' 	public stwk_work_stop_date
' 	public stwk_income_stop_date
' 	public stwk_verification
' 	public stwk_refused_employment
' 	public stwk_vol_quit
' 	public stwk_refused_employment_date
' 	public stwk_cash_good_cause_yn
' 	public stwk_grh_good_cause_yn
' 	public stwk_snap_good_cause_yn
' 	public stwk_snap_pwe
' 	public stwk_ma_epd_extension
' 	public stwk_summary
'
' 	public fmed_exists
' 	public fmed_miles
' 	public fmed_rate
' 	public fmed_milage_expense
' 	public fmed_page()
' 	public fmed_row()
' 	public fmed_type()
' 	public fmed_verif()
' 	public fmed_ref()
' 	public fmed_catgry()
' 	public fmed_begin()
' 	public fmed_end()
' 	public fmed_expense()
' 	public fmed_notes()
'
' 	public pded_exists
' 	public pded_guardian_fee
' 	public pded_rep_payee_fee
' 	public pded_shel_spec_need
'
' 	public diet_exists
' 	public diet_mf_type_one
' 	public diet_mf_verif_one
' 	public diet_mf_type_two
' 	public diet_mf_verif_two
' 	public diet_msa_type_one
' 	public diet_msa_verif_one
' 	public diet_msa_type_two
' 	public diet_msa_verif_two
' 	public diet_msa_type_three
' 	public diet_msa_verif_three
' 	public diet_msa_type_four
' 	public diet_msa_verif_four
' 	public diet_msa_type_five
' 	public diet_msa_verif_five
' 	public diet_msa_type_six
' 	public diet_msa_verif_six
' 	public diet_msa_type_seven
' 	public diet_msa_verif_seven
' 	public diet_msa_type_eight
' 	public diet_msa_verif_eight
'
' 	public checkbox_one
' 	public checkbox_two
' 	public checkbox_three
' 	public checkbox_four
'
' 	public detail_one
' 	public detail_two
' 	public detail_three
' 	public detail_four
'
' 	public button_one
' 	public button_two
' 	public button_three
' 	public button_four
'
' 	public clt_has_cs_income
' 	public clt_cs_counted
' 	public cs_paid_to
' 	public clt_has_ss_income
' 	public clt_has_BUSI
' 	public clt_has_JOBS
'
' 	public snap_req_checkbox
' 	public cash_req_checkbox
' 	public emer_req_checkbox
' 	public grh_req_checkbox
' 	public hc_req_checkbox
' 	public none_req_checkbox
' 	public client_verification
' 	public client_verification_details
' 	public client_notes
'
' 	public property get full_name_const
' 		full_name_const = first_name & " " & last_name
' 	end property
'
' 	Public sub define_the_member()
'
' 		pare_child_ref_nbr = array("")
' 		pare_child_name = array("")
' 		pare_child_member_index = array("")
' 		pare_relationship_type = array("")
' 		pare_verification = array("")
'
' 		intend_to_reside_in_mn = "Yes"
'
' 		Call navigate_to_MAXIS_screen("STAT", "MEMB")		'===============================================================================================
' 		EMWriteScreen ref_number, 20, 76
' 		transmit
'
' 		EMReadScreen access_denied_check, 13, 24, 2         'Sometimes MEMB gets this access denied issue and we have to work around it.
' 		If access_denied_check = "ACCESS DENIED" Then
' 			PF10
' 			last_name = "UNABLE TO FIND"
' 			first_name = "Access Denied"
' 			mid_initial = ""
' 			access_denied = TRUE
' 		Else
' 			access_denied = FALSE
' 			EMReadscreen last_name, 25, 6, 30
' 			EMReadscreen first_name, 12, 6, 63
' 			EMReadscreen mid_initial, 1, 6, 79
' 			EMReadScreen age, 3, 8, 76
'
' 			EMReadScreen date_of_birth, 10, 8, 42
' 			EMReadScreen ssn, 11, 7, 42
' 			EMReadScreen ssn_verif, 1, 7, 68
' 			EMReadScreen birthdate_verif, 2, 8, 68
' 			EMReadScreen gender, 1, 9, 42
' 			EMReadScreen race, 30, 17, 42
' 			EMReadScreen spoken_lang, 20, 12, 42
' 			EMReadScreen written_lang, 29, 13, 42
' 			EMReadScreen interpreter, 1, 14, 68
' 			EMReadScreen alias_yn, 1, 15, 42
' 			EMReadScreen ethnicity_yn, 1, 16, 68
'
' 			age = trim(age)
' 			If age = "" Then age = 0
' 			age = age * 1
' 			last_name = trim(replace(last_name, "_", ""))
' 			first_name = trim(replace(first_name, "_", ""))
' 			mid_initial = replace(mid_initial, "_", "")
' 			EMReadScreen id_verif, 2, 9, 68
'
' 			EMReadScreen rel_to_applcnt, 2, 10, 42              'reading the relationship from MEMB'
' 			If rel_to_applcnt = "01" Then rel_to_applcnt = "Self"
' 			If rel_to_applcnt = "02" Then rel_to_applcnt = "Spouse"
' 			If rel_to_applcnt = "03" Then rel_to_applcnt = "Child"
' 			If rel_to_applcnt = "04" Then rel_to_applcnt = "Parent"
' 			If rel_to_applcnt = "05" Then rel_to_applcnt = "Sibling"
' 			If rel_to_applcnt = "06" Then rel_to_applcnt = "Step Sibling"
' 			If rel_to_applcnt = "08" Then rel_to_applcnt = "Step Child"
' 			If rel_to_applcnt = "09" Then rel_to_applcnt = "Step Parent"
' 			If rel_to_applcnt = "10" Then rel_to_applcnt = "Aunt"
' 			If rel_to_applcnt = "11" Then rel_to_applcnt = "Uncle"
' 			If rel_to_applcnt = "12" Then rel_to_applcnt = "Niece"
' 			If rel_to_applcnt = "13" Then rel_to_applcnt = "Nephew"
' 			If rel_to_applcnt = "14" Then rel_to_applcnt = "Cousin"
' 			If rel_to_applcnt = "15" Then rel_to_applcnt = "Grandparent"
' 			If rel_to_applcnt = "16" Then rel_to_applcnt = "Grandchild"
' 			If rel_to_applcnt = "17" Then rel_to_applcnt = "Other Relative"
' 			If rel_to_applcnt = "18" Then rel_to_applcnt = "Legal Guardian"
' 			If rel_to_applcnt = "24" Then rel_to_applcnt = "Not Related"
' 			If rel_to_applcnt = "25" Then rel_to_applcnt = "Live-in Attendant"
' 			If rel_to_applcnt = "27" Then rel_to_applcnt = "Unknown"
'
' 			If id_verif = "BC" Then id_verif = "BC - Birth Certificate"
' 			If id_verif = "RE" Then id_verif = "RE - Religious Record"
' 			If id_verif = "DL" Then id_verif = "DL - Drivers License/ST ID"
' 			If id_verif = "DV" Then id_verif = "DV - Divorce Decree"
' 			If id_verif = "AL" Then id_verif = "AL - Alien Card"
' 			If id_verif = "AD" Then id_verif = "AD - Arrival//Depart"
' 			If id_verif = "DR" Then id_verif = "DR - Doctor Stmt"
' 			If id_verif = "PV" Then id_verif = "PV - Passport/Visa"
' 			If id_verif = "OT" Then id_verif = "OT - Other Document"
' 			If id_verif = "NO" Then id_verif = "NO - No Veer Prvd"
'
' 			If age > 18 then
' 				cash_minor = FALSE
' 			Else
' 				cash_minor = TRUE
' 			End If
' 			If age > 21 then
' 				snap_minor = FALSE
' 			Else
' 				snap_minor = TRUE
' 			End If
'
' 			date_of_birth = replace(date_of_birth, " ", "/")
' 			If birthdate_verif = "BC" Then birthdate_verif = "BC - Birth Certificate"
' 			If birthdate_verif = "RE" Then birthdate_verif = "RE - Religious Record"
' 			If birthdate_verif = "DL" Then birthdate_verif = "DL - Drivers License/State ID"
' 			If birthdate_verif = "DV" Then birthdate_verif = "DV - Divorce Decree"
' 			If birthdate_verif = "AL" Then birthdate_verif = "AL - Alien Card"
' 			If birthdate_verif = "DR" Then birthdate_verif = "DR - Doctor Statement"
' 			If birthdate_verif = "OT" Then birthdate_verif = "OT - Other Document"
' 			If birthdate_verif = "PV" Then birthdate_verif = "PV - Passport/Visa"
' 			If birthdate_verif = "NO" Then birthdate_verif = "NO - No Verif Provided"
'
' 			ssn = replace(ssn, " ", "-")
' 			if ssn = "___-__-____" Then ssn = ""
' 			If ssn_verif = "A" THen ssn_verif = "A - SSN Applied For"
' 			If ssn_verif = "P" THen ssn_verif = "P - SSN Provided, verif Pending"
' 			If ssn_verif = "N" THen ssn_verif = "N - SSN Not Provided"
' 			If ssn_verif = "V" THen ssn_verif = "V - SSN Verified via Interface"
'
' 			If gender = "M" Then gender = "Male"
' 			If gender = "F" Then gender = "Female"
'
' 			race = trim(race)
'
' 			spoken_lang = replace(replace(spoken_lang, "_", ""), "  ", " - ")
' 			written_lang = trim(replace(replace(replace(written_lang, "_", ""), "  ", " - "), "(HRF)", ""))
'
' 			clt_has_cs_income = FALSE
' 			clt_has_ss_income = FALSE
' 			clt_has_BUSI = FALSE
' 			clt_has_JOBS = FALSE
' 		End If
'
' 		If access_denied = FALSE Then
' 			Call navigate_to_MAXIS_screen("STAT", "MEMI")		'===============================================================================================
' 			EMWriteScreen ref_number, 20, 76
' 			transmit
'
' 			EMReadScreen marital_status, 1, 7, 40
' 			EMReadScreen spouse_ref, 2, 9, 49
' 			EMReadScreen spouse_name, 40, 9, 52
' 			EMReadScreen last_grade_completed, 2, 10, 49
' 			EMReadScreen citizen, 1, 11, 49
' 			EMReadScreen other_st_FS_end_date, 8, 13, 49
' 			EMReadScreen in_mn_12_mo, 1, 14, 49
' 			EMReadScreen residence_verif, 1, 14, 78
' 			EMReadScreen mn_entry_date, 8, 15, 49
' 			EMReadScreen former_state, 2, 15, 78
'
' 			If marital_status = "N" Then marital_status = "N - Never Married"
' 			If marital_status = "M" Then marital_status = "M - Married Living with Spouse"
' 			If marital_status = "S" Then marital_status = "S - Married Living Apart"
' 			If marital_status = "L" Then marital_status = "L - Legally Seperated"
' 			If marital_status = "D" Then marital_status = "D - Divorced"
' 			If marital_status = "W" Then marital_status = "W - Widowed"
' 			If spouse_ref = "__" Then spouse_ref = ""
' 			spouse_name = trim(spouse_name)
'
' 			If last_grade_completed = "00" Then last_grade_completed = "Not Attended or Pre-Grade 1 - 00"
' 			If last_grade_completed = "12" Then last_grade_completed = "High School Diploma or GED - 12"
' 			If last_grade_completed = "13" Then last_grade_completed = "Some Post Sec Education - 13"
' 			If last_grade_completed = "14" Then last_grade_completed = "High School Plus Certiificate - 14"
' 			If last_grade_completed = "15" Then last_grade_completed = "Four Year Degree - 15"
' 			If last_grade_completed = "16" Then last_grade_completed = "Grad Degree - 16"
' 			If len(last_grade_completed) = 2 Then last_grade_completed = "Grade " & last_grade_completed
' 			If citizen = "Y" Then citizen = "Yes"
' 			If citizen = "N" Then citizen = "No"
'
' 			other_st_FS_end_date = replace(other_st_FS_end_date, " ", "/")
' 			If other_st_FS_end_date = "__/__/__" Then other_st_FS_end_date = ""
' 			If in_mn_12_mo = "Y" Then in_mn_12_mo = "Yes"
' 			If in_mn_12_mo = "N" Then in_mn_12_mo = "No"
' 			If residence_verif = "1" Then residence_verif = "1 - Rent Receipt"
' 			If residence_verif = "2" Then residence_verif = "2 - Landlord's Statement"
' 			If residence_verif = "3" Then residence_verif = "3 - Utility Bill"
' 			If residence_verif = "4" Then residence_verif = "4 - Other"
' 			If residence_verif = "N" Then residence_verif = "N - Verif Not Provided"
' 			mn_entry_date = replace(mn_entry_date, " ", "/")
' 			If mn_entry_date = "__/__/__" Then mn_entry_date = ""
' 			If former_state = "__" Then former_state = ""
'
'
' 			Call navigate_to_MAXIS_screen("STAT", "IMIG")		'===============================================================================================
' 			EMWriteScreen ref_number, 20, 76
' 			transmit
'
' 			EMreadScreen imig_version, 1, 2, 73
' 			If imig_version = "0" Then imig_exists = FALSE
' 			If imig_version = "1" Then imig_exists = TRUE
'
' 			If imig_exists = TRUE Then
' 				EMReadScreen imig_status_code, 2, 6, 45
' 				EMReadScreen imig_status_desc, 32, 6, 48
' 				EMReadScreen us_entry_date, 10, 7, 45
' 				EMReadScreen imig_status_date, 10, 7, 71
' 				EMReadScreen imig_status_verif, 2, 8, 45
' 				EMReadScreen lpr_adj_from, 40, 9, 45
' 				EMReadScreen nationality, 2, 10, 45
' 				EMReadScreen alien_id_nbr, 10, 10, 71
'
' 				imig_status_desc = trim(imig_status_desc)
' 				imig_status = imig_status_code & " - " & imig_status_desc
' 				us_entry_date = replace(us_entry_date, " ", "/")
' 				imig_status_date = replace(imig_status_date, " ", "/")
'
' 				If imig_status_verif = "S1" Then imig_status_verif = "S1 - SAVE Primary"
' 				If imig_status_verif = "S2" Then imig_status_verif = "S2 - SAVE Secondary"
' 				If imig_status_verif = "AL" Then imig_status_verif = "AL - Alien Card"
' 				If imig_status_verif = "PV" Then imig_status_verif = "PV - Passport/Visa"
' 				If imig_status_verif = "RE" Then imig_status_verif = "RE - Re-Entry Permit"
' 				If imig_status_verif = "IM" Then imig_status_verif = "IN - INS Correspondence"
' 				If imig_status_verif = "OT" Then imig_status_verif = "OT - Other Document"
' 				If imig_status_verif = "NO" Then imig_status_verif = "NO - No Verif Provided"
'
' 				lpr_adj_from = trim(lpr_adj_from)
'
' 				If nationality = "AA" Then nationality = "AA - Amerasian"
' 				If nationality = "EH" Then nationality = "EH - Ethnic Chinese"
' 				If nationality = "EL" Then nationality = "EL - Ethnic Lao"
' 				If nationality = "HG" Then nationality = "HG - Hmong"
' 				If nationality = "KD" Then nationality = "KD - Kurd"
' 				If nationality = "SJ" Then nationality = "SJ - Soviet Jew"
' 				If nationality = "TT" Then nationality = "TT - Tinh"
' 				If nationality = "AF" Then nationality = "AF - Afghanistan"
' 				If nationality = "BK" Then nationality = "BK - Bosnia"
' 				If nationality = "CB" Then nationality = "CB - Cambodia"
' 				If nationality = "CH" Then nationality = "CH - China, Mainland"
' 				If nationality = "CU" Then nationality = "CU - Cuba"
' 				If nationality = "ES" Then nationality = "ES - El Salvador"
' 				If nationality = "ER" Then nationality = "ER - Eritrea"
' 				If nationality = "ET" Then nationality = "ET - Ethiopia"
' 				If nationality = "GT" Then nationality = "GT - Guatemala"
' 				If nationality = "HA" Then nationality = "HA - Haiti"
' 				If nationality = "HO" Then nationality = "HO - Honduras"
' 				If nationality = "IR" Then nationality = "IR - Iran"
' 				If nationality = "IZ" Then nationality = "IZ - Iraq"
' 				If nationality = "LI" Then nationality = "LI - Liberia"
' 				If nationality = "MC" Then nationality = "MC - Micronesia"
' 				If nationality = "MI" Then nationality = "MI - Marshall Islands"
' 				If nationality = "MX" Then nationality = "MX - Mexico"
' 				If nationality = "WA" Then nationality = "WA - Namibia (SW Africa)"
' 				If nationality = "PK" Then nationality = "PK - Pakistan"
' 				If nationality = "RP" Then nationality = "RP - Philippines"
' 				If nationality = "PL" Then nationality = "PL - Poland"
' 				If nationality = "RO" Then nationality = "RO - Romania"
' 				If nationality = "RS" Then nationality = "RS - Russia"
' 				If nationality = "SO" Then nationality = "SO - Somalia"
' 				If nationality = "SF" Then nationality = "SF - South Africa"
' 				If nationality = "TH" Then nationality = "TH - Thailand"
' 				If nationality = "VM" Then nationality = "VM - Vietnam"
' 				If nationality = "OT" Then nationality = "OT - All Others"
'
' 				imig_q_2_required = TRUE
' 				imig_q_4_required = TRUE
' 				imig_q_5_required = TRUE
'
' 			End If
'
' 			Call navigate_to_MAXIS_screen("STAT", "SPON")		'===============================================================================================
' 			EMWriteScreen ref_number, 20, 76
' 			transmit
'
' 			EMreadScreen spon_version, 1, 2, 73
' 			If spon_version = "0" Then spon_exists = FALSE
' 			If spon_version = "1" Then spon_exists = TRUE
' 			clt_has_sponsor = "No"
'
' 			If spon_exists = TRUE Then
' 				clt_has_sponsor = "Yes"
' 				' new_spon_name			=
' 				' new_spon_street			=
' 				' new_spon_city			=
' 				' new_spon_state			=
' 				' new_spon_zip			=
' 				' new_spon_phone			=
' 				' new_spon_gross_income	=
' 				' new_spon_income_freq	=
' 				' new_spon_spouse_name	=
' 				' new_spon_spouse_income	=
' 				' new_spon_spouse_income_freq =
'
'
' 			End If
' 			' public spon_exists
'
' 			Call navigate_to_MAXIS_screen("STAT", "REMO")		'===============================================================================================
' 			EMWriteScreen ref_number, 20, 76
' 			transmit
'
' 			EMreadScreen remo_version, 1, 2, 73
' 			If remo_version = "0" Then remo_exists = FALSE
' 			If remo_version = "1" Then remo_exists = TRUE
'
' 			If remo_exists = TRUE Then
' 				EMReadScreen left_hh_date, 8, 8, 53
' 				EMReadScreen left_hh_reason, 2, 8, 71
' 				EMReadScreen left_hh_expected_return_date, 8, 13, 53
' 				EMReadScreen left_hh_expected_return_verif, 2, 13, 71
' 				EMReadScreen left_hh_actual_return_date, 8, 14, 53
' 				EMReadScreen left_hh_HC_temp_out_of_state, 1, 16, 64
' 				EMReadScreen left_hh_date_reported, 8, 17, 64
'
' 				left_hh_date = replace(left_hh_date, " ", "/")
' 				If left_hh_date = "__/__/__" Then left_hh_date = ""
'
' 				If left_hh_reason = "01" Then left_hh_reason = "01 - Death"
' 				If left_hh_reason = "02" Then left_hh_reason = "02 - Moved out of Household"
' 				If left_hh_reason = "03" Then left_hh_reason = "03 - Institutional Placement"
' 				If left_hh_reason = "04" Then left_hh_reason = "04 - IV-E Foster Care Placement"
' 				If left_hh_reason = "05" Then left_hh_reason = "05 - Non IV-E Foster Care Placement"
' 				If left_hh_reason = "06" Then left_hh_reason = "06 - Illness"
' 				If left_hh_reason = "07" Then left_hh_reason = "07 - Vacation or Visit"
' 				If left_hh_reason = "08" Then left_hh_reason = "08 - Runaway"
' 				If left_hh_reason = "09" Then left_hh_reason = "09 - Away for Education"
' 				If left_hh_reason = "10" Then left_hh_reason = "10 - Relative Ill/Deceased"
' 				If left_hh_reason = "11" Then left_hh_reason = "11 - Training of Employment Search"
' 				If left_hh_reason = "12" Then left_hh_reason = "12 - Incarceration"
' 				If left_hh_reason = "13" Then left_hh_reason = "13 - Other Allowed Return before 10th"
' 				If left_hh_reason = "14" Then left_hh_reason = "14 - Non-Allowed Absent Cash"
' 				If left_hh_reason = "15" Then left_hh_reason = "15 - Military Service"
' 				If left_hh_reason = "16" Then left_hh_reason = "16 - Other"
' 				If left_hh_reason = "__" Then left_hh_reason = ""
'
' 				left_hh_expected_return_date = replace(left_hh_expected_return_date, " ", "/")
' 				If left_hh_expected_return_date = "__/__/__" Then left_hh_expected_return_date = ""
'
' 				If left_hh_expected_return_verif = "01" Then left_hh_expected_return_verif = "01 - Social Worker Statement"
' 				If left_hh_expected_return_verif = "02" Then left_hh_expected_return_verif = "02 - Court Papers"
' 				If left_hh_expected_return_verif = "03" Then left_hh_expected_return_verif = "03 - Doctor Statement"
' 				If left_hh_expected_return_verif = "04" Then left_hh_expected_return_verif = "04 - Other Document"
' 				If left_hh_expected_return_verif = "__" Then left_hh_expected_return_verif = ""
'
' 				left_hh_actual_return_date = replace(left_hh_actual_return_date, " ", "/")
' 				If left_hh_actual_return_date = "__/__/__" Then left_hh_actual_return_date = ""
'
' 				If left_hh_HC_temp_out_of_state = "_" Then left_hh_HC_temp_out_of_state = ""
'
' 				left_hh_date_reported = replace(left_hh_date_reported, " ", "/")
' 				If left_hh_date_reported = "__/__/__" Then left_hh_date_reported = ""
'
' 				If IsDate(left_hh_date) = TRUE Then
' 					If DateDiff("m", left_hh_date, date) >= 12 Then
' 						left_hh_12_months_or_more = TRUE
' 					Else
' 						left_hh_12_months_or_more = FALSE
' 					End If
' 				End If
' 			End If
'
' 			Call navigate_to_MAXIS_screen("STAT", "ADME")		'===============================================================================================
' 			EMWriteScreen ref_number, 20, 76
' 			transmit
'
' 			EMreadScreen adme_version, 1, 2, 73
' 			If adme_version = "0" Then adme_exists = FALSE
' 			If adme_version = "1" Then adme_exists = TRUE
'
' 			If adme_exists = TRUE Then
' 				EMReadScreen adme_arrival_date, 8, 7, 38
' 				EMReadScreen adme_cash_date, 8, 12, 38
' 				EMReadScreen adme_emer_date, 8, 14, 38
' 				EMReadScreen adme_snap_date, 8, 16, 38
'
' 				adme_arrival_date = trim(adme_arrival_date)
' 				If adme_arrival_date = "////////" Then adme_arrival_date = ""
'
' 				adme_cash_date = replace(adme_cash_date, " ", "/")
' 				If adme_cash_date = "__/__/__" Then adme_cash_date = ""
'
' 				adme_emer_date = replace(adme_emer_date, " ", "/")
' 				If adme_emer_date = "__/__/__" Then adme_emer_date = ""
'
' 				adme_snap_date = replace(adme_snap_date, " ", "/")
' 				If adme_snap_date = "__/__/__" Then adme_snap_date = ""
'
' 				adme_within_12_months = FALSE
' 				If IsDate(adme_arrival_date) = TRUE Then
' 					If DateDiff("m", adme_arrival_date, date) < 13 Then adme_within_12_months = TRUE
' 				End If
' 			End If
'
'
' 			Call navigate_to_MAXIS_screen("STAT", "COEX")		'===============================================================================================
' 			EMWriteScreen ref_number, 20, 76
' 			transmit
'
' 			EMreadScreen coex_version, 1, 2, 73
' 			If coex_version = "0" Then coex_exists = FALSE
' 			If coex_version = "1" Then coex_exists = TRUE
'
' 			If coex_exists = TRUE Then
' 				EMReadScreen coex_support_verif, 1, 10, 36
' 				EMReadScreen coex_support_retro_amount, 8, 10, 45
' 				EMReadScreen coex_support_prosp_amount, 8, 10, 63
'
' 				EMReadScreen coex_alimony_verif, 1, 11, 36
' 				EMReadScreen coex_alimony_retro_amount, 8, 11, 45
' 				EMReadScreen coex_alimony_prosp_amount, 8, 11, 63
'
' 				EMReadScreen coex_tax_dep_verif, 1, 12, 36
' 				EMReadScreen coex_tax_dep_retro_amount, 8, 12, 45
' 				EMReadScreen coex_tax_dep_prosp_amount, 8, 12, 63
'
' 				EMReadScreen coex_other_verif, 1, 13, 36
' 				EMReadScreen coex_other_retro_amount, 8, 13, 45
' 				EMReadScreen coex_other_prosp_amount, 8, 13, 63
'
' 				EMReadScreen coex_total_retro_amount, 8, 15, 45
' 				EMReadScreen coex_total_prosp_amount, 8, 15, 63
'
' 				EMReadScreen coex_change_in_financial_circumstances, 1, 17, 61
'
' 				EMWriteScreen "X", 18, 44
' 				transmit
'
' 				EMReadScreen coex_support_hc_est_amount, 8, 6, 38
' 				EMReadScreen coex_alimony_hc_est_amount, 8, 7, 38
' 				EMReadScreen coex_tax_dep_hc_est_amount, 8, 8, 38
' 				EMReadScreen coex_other_hc_est_amount, 8, 9, 38
' 				EMReadScreen coex_total_hc_est_amount, 8, 11, 38
'
' 				PF3
'
' 			End If
'
' 			Call navigate_to_MAXIS_screen("STAT", "DISA")		'===============================================================================================
' 			EMWriteScreen ref_number, 20, 76
' 			transmit
'
' 			EMreadScreen disa_version, 1, 2, 73
' 			If disa_version = "0" Then disa_exists = FALSE
' 			If disa_version = "1" Then disa_exists = TRUE
'
' 			If disa_exists = TRUE Then
' 				EMReadScreen disa_begin_date, 10, 6, 47
' 				EMReadScreen disa_end_date, 10, 6, 69
' 				EMReadScreen disa_cert_begin_date, 10, 7, 47
' 				EMReadScreen disa_cert_end_date, 10, 7, 69
' 				EMReadScreen cash_disa_status, 2, 11, 59
' 				EMReadScreen cash_disa_verif, 1, 11, 69
' 				EMReadScreen fs_disa_status, 2, 12, 59
' 				EMReadScreen fs_disa_verif, 1, 12, 69
' 				EMReadScreen hc_disa_status, 2, 13, 59
' 				EMReadScreen hc_disa_verif, 1, 13, 69
' 				EMReadScreen disa_waiver, 1, 14, 59
' 				EMReadScreen disa_1619, 1, 16, 59
'
' 				disa_begin_date = replace(disa_begin_date, " ", "/")
' 				If disa_begin_date = "__/__/____" Then disa_begin_date = ""
' 				disa_end_date = replace(disa_end_date, " ", "/")
' 				If disa_end_date = "__/__/____" Then disa_end_date = ""
' 				disa_cert_begin_date = replace(disa_cert_begin_date, " ", "/")
' 				If disa_cert_begin_date = "__/__/____" Then disa_cert_begin_date = ""
' 				disa_cert_end_date = replace(disa_cert_end_date, " ", "/")
' 				If disa_cert_end_date = "__/__/____" Then disa_cert_end_date = ""
'
' 				If hc_disa_verif = "1" OR fs_disa_verif = "1" OR cash_disa_status = "1" Then disa_detail = "DISA based on Doctor's Statement"
' 				If hc_disa_verif = "2" OR fs_disa_verif = "2" OR cash_disa_status = "2" Then disa_detail = "SMRT Certified Disability"
' 				If hc_disa_verif = "3" OR fs_disa_verif = "3" OR cash_disa_status = "3" Then disa_detail = "SSA Certified Disability"
' 				If cash_disa_status = "7" Then disa_detail = "Disability based on Professional Statement of Need"
'
' 				If cash_disa_status = "01" Then cash_disa_status = "01 - RSDI Only Disability"
' 				If cash_disa_status = "02" Then cash_disa_status = "02 - RSDI Only Blindness"
' 				If cash_disa_status = "03" Then cash_disa_status = "03 - SSI, SSI/RSDI Disability"
' 				If cash_disa_status = "04" Then cash_disa_status = "04 - SSI, SSI/RSDI Blindness"
' 				If cash_disa_status = "06" Then cash_disa_status = "06 - SMRT/SSA Pend"
' 				If cash_disa_status = "08" Then cash_disa_status = "08 - SMRT Certified Blindness"
' 				If cash_disa_status = "09" Then cash_disa_status = "09 - Ill/Incapacity"
' 				If cash_disa_status = "10" Then cash_disa_status = "10 - SMRT Certified Disability"
' 				If cash_disa_status = "__" Then cash_disa_status = ""
' 				If cash_disa_verif = "1" Then cash_disa_verif = "1 - DHS 161/Dr Stmt"
' 				If cash_disa_verif = "2" Then cash_disa_verif = "2 - SMRT Certified"
' 				If cash_disa_verif = "3" Then cash_disa_verif = "3 - Certified for RSDI or SSI"
' 				If cash_disa_verif = "6" Then cash_disa_verif = "6 - Other Document"
' 				If cash_disa_verif = "7" Then cash_disa_verif = "7 - Professional Stmt of Need"
' 				If cash_disa_verif = "N" Then cash_disa_verif = "N - No Verif Provided"
'
' 				If fs_disa_status = "01" Then fs_disa_status = "01 - RSDI Only Disability"
' 				If fs_disa_status = "02" Then fs_disa_status = "02 - RSDI Only Blindness"
' 				If fs_disa_status = "03" Then fs_disa_status = "03 - SSI, SSI/RSDI Disability"
' 				If fs_disa_status = "04" Then fs_disa_status = "04 - SSI, SSI/RSDI Blindness"
' 				If fs_disa_status = "08" Then fs_disa_status = "08 - SMRT Certified Blindness"
' 				If fs_disa_status = "09" Then fs_disa_status = "09 - Ill/Incapacity"
' 				If fs_disa_status = "10" Then fs_disa_status = "10 - SMRT Certified Disability"
' 				If fs_disa_status = "11" Then fs_disa_status = "11 - VA Determined Pd - 100% Disa"
' 				If fs_disa_status = "12" Then fs_disa_status = "12 - VA (Other Accept Disa)"
' 				If fs_disa_status = "13" Then fs_disa_status = "13 - Certified RR Retirement Disa"
' 				If fs_disa_status = "14" Then fs_disa_status = "14 - Other Govt Permanent Disa"
' 				If fs_disa_status = "15" Then fs_disa_status = "15 - Disability from MINE List"
' 				If fs_disa_status = "16" Then fs_disa_status = "16 - Unable to Prepare Purch Own Meal"
' 				If fs_disa_status = "__" Then fs_disa_status = ""
' 				If fs_disa_verif = "1" Then fs_disa_verif = "1 - DHS 161/Dr Stmt"
' 				If fs_disa_verif = "2" Then fs_disa_verif = "2 - SMRT Certified"
' 				If fs_disa_verif = "3" Then fs_disa_verif = "3 - Certified for RSDI or SSI"
' 				If fs_disa_verif = "4" Then fs_disa_verif = "4 - Receipt of HC for Disa/Blind"
' 				If fs_disa_verif = "5" Then fs_disa_verif = "5 - Work Judgement"
' 				If fs_disa_verif = "6" Then fs_disa_verif = "6 - Other Document"
' 				If fs_disa_verif = "7" Then fs_disa_verif = "7 - Out of State Verif Pending"
' 				If fs_disa_verif = "N" Then fs_disa_verif = "N - No Verif Provided"
'
' 				If hc_disa_status = "01" Then hc_disa_status = "01 - RSDI Only Disability"
' 				If hc_disa_status = "02" Then hc_disa_status = "02 - RSDI Only Blindness"
' 				If hc_disa_status = "03" Then hc_disa_status = "03 - SSI, SSI/RSDI Disability"
' 				If hc_disa_status = "04" Then hc_disa_status = "04 - SSI, SSI/RSDI Blindness"
' 				If hc_disa_status = "06" Then hc_disa_status = "06 - SMRT Pend or SSA Pend"
' 				If hc_disa_status = "08" Then hc_disa_status = "08 - Certified Blind"
' 				If hc_disa_status = "10" Then hc_disa_status = "10 - Certified Disabled"
' 				If hc_disa_status = "11" Then hc_disa_status = "11 - Special Category - Disabled Child"
' 				If hc_disa_status = "20" Then hc_disa_status = "20 - TEFRA - Disabled"
' 				If hc_disa_status = "21" Then hc_disa_status = "21 - TEFRA - Blind"
' 				If hc_disa_status = "22" Then hc_disa_status = "22 - MA-EPD"
' 				If hc_disa_status = "23" Then hc_disa_status = "23 - MA/Waiver"
' 				If hc_disa_status = "24" Then hc_disa_status = "24 - SSA/SMRT Appeal Pending"
' 				If hc_disa_status = "26" Then hc_disa_status = "26 - SSA/SMRT Disa Deny"
' 				If hc_disa_status = "__" Then hc_disa_status = ""
' 				If hc_disa_verif = "1" Then hc_disa_verif = "1 - DHS 161/Dr Stmt"
' 				If hc_disa_verif = "2" Then hc_disa_verif = "2 - SMRT Certified"
' 				If hc_disa_verif = "3" Then hc_disa_verif = "3 - Certified for RSDI or SSI"
' 				If hc_disa_verif = "6" Then hc_disa_verif = "6 - Other Document"
' 				If hc_disa_verif = "7" Then hc_disa_verif = "7 - Case Manager Determination"
' 				If hc_disa_verif = "8" Then hc_disa_verif = "8 - LTC Consult Services"
' 				If hc_disa_verif = "N" Then hc_disa_verif = "N - No Verif Provided"
'
' 				If disa_waiver = "F" Then disa_waiver = "F - LTC CADI Conversion"
' 				If disa_waiver = "G" Then disa_waiver = "G - LTC CADI DIversion"
' 				If disa_waiver = "H" Then disa_waiver = "H - LTC CAC Conversion"
' 				If disa_waiver = "I" Then disa_waiver = "I - LTC CAC Diversion"
' 				If disa_waiver = "J" Then disa_waiver = "J - LTC EW Conversion"
' 				If disa_waiver = "K" Then disa_waiver = "K - LTC EW Diversion"
' 				If disa_waiver = "L" Then disa_waiver = "L - LTC TBI NF Conversion"
' 				If disa_waiver = "M" Then disa_waiver = "M - LTC TBI NF Diversion"
' 				If disa_waiver = "P" Then disa_waiver = "P - LTC TBI NB Conversion"
' 				If disa_waiver = "Q" Then disa_waiver = "Q - LTC TBI NB Diversion"
' 				If disa_waiver = "R" Then disa_waiver = "R - DD Conversion"
' 				If disa_waiver = "S" Then disa_waiver = "S - DD Conversion"
' 				If disa_waiver = "Y" Then disa_waiver = "Y - CSG Conversion"
' 				If disa_waiver = "_" Then disa_waiver = ""
'
' 				If disa_1619 = "A" Then disa_1619 = "A - 1619A Status"
' 				If disa_1619 = "B" Then disa_1619 = "B - 1619B Status"
' 				If disa_1619 = "N" Then disa_1619 = "N - No 1619 Status"
' 				If disa_1619 = "T" Then disa_1619 = "T - 1619 Status Terminated"
' 				If disa_1619 = "_" Then disa_1619 = ""
' 			End If
'
' 			Call navigate_to_MAXIS_screen("STAT", "WREG")		'===============================================================================================
' 			EMWriteScreen ref_number, 20, 76
' 			transmit
'
' 			EMreadScreen wreg_version, 1, 2, 73
' 			If wreg_version = "0" Then wreg_exists = FALSE
' 			If wreg_version = "1" Then wreg_exists = TRUE
'
' 			If wreg_exists = TRUE Then
' 				EMReadScreen wreg_pwe, 1, 6, 68
'
' 				If wreg_pwe = "Y" Then fs_pwe = "Yes"
' 				If wreg_pwe = "N" OR wreg_pwe = "_" Then fs_pwe = "No"
' 			End If
'
'
' 			Call navigate_to_MAXIS_screen("STAT", "SCHL")		'===============================================================================================
' 			EMWriteScreen ref_number, 20, 76
' 			transmit
'
' 			EMreadScreen schl_version, 1, 2, 73
' 			If schl_version = "0" Then schl_exists = FALSE
' 			If schl_version = "1" Then schl_exists = TRUE
'
' 			If schl_exists = TRUE Then
' 				EMReadScreen schl_status, 1, 6, 40
' 				EMReadScreen schl_verif, 2, 6, 63
' 				EMReadScreen schl_type, 2, 7, 40
' 				EMReadScreen school_district, 4, 8, 40
' 				EMReadScreen schl_start_date, 8, 10, 63
' 				EMReadScreen schl_grad_date, 5, 11, 63
' 				EMReadScreen schl_grad_verif, 2, 12, 63
' 				EMReadScreen schl_fund, 1, 14, 63
' 				EMReadScreen schl_elig, 2, 16, 63
' 				EMReadScreen schl_higher_ed_yn, 1, 18, 63
'
' 				If schl_status = "F" Then school_status = "Fulltime"
' 				If schl_status = "H" Then school_status = "Halftime"
' 				If schl_status = "L" Then school_status = "Less than Half "
' 				If schl_status = "N" Then school_status = "Not Attending"
'
' 				If schl_verif = "SC" Then school_verif = "SC - School Statement"
' 				If schl_verif = "OT" Then school_verif = "OT - Other Document"
' 				If schl_verif = "NO" Then school_verif = "NO - No Verif Provided"
' 				If schl_verif = "__" Then school_verif = "Blank"
'
' 				If schl_type = "01" Then school_type = "01 - Preschool - 6"
' 				If schl_type = "11" Then school_type = "11 - 7 - 8"
' 				If schl_type = "02" Then school_type = "02 - 9 - 12"
' 				If schl_type = "03" Then school_type = "03 - GED Or Equiv"
' 				If schl_type = "06" Then school_type = "06 - Child, Not In School"
' 				If schl_type = "07" Then school_type = "07 - Individual Ed Plan/IEP"
' 				If schl_type = "08" Then school_type = "08 - Post-Sec Not Grad Student"
' 				If schl_type = "09" Then school_type = "09 - Post-Sec Grad Student"
' 				If schl_type = "10" Then school_type = "10 - Post-Sec Tech Schl"
' 				If schl_type = "12" Then school_type = "11 - Adult Basic Ed (ABE)"
' 				If schl_type = "13" Then school_type = "13 - English As A 2nd Language"
'
' 				If school_district = "____" Then school_district = ""
'
' 				kinder_start_date = replace(schl_start_date, " ", "/")
' 				If kinder_start_date = "__/__/__" Then kinder_start_date = ""
'
' 				grad_date = replace(schl_grad_date, " ", "/")
' 				If grad_date = "__/__" Then grad_date = ""
'
' 				If schl_grad_verif = "SC" Then grad_date_verif = "SC - School Statement"
' 				If schl_grad_verif = "OT" Then grad_date_verif = "OT - Other Document"
' 				If schl_grad_verif = "NO" Then grad_date_verif = "NO - No Verif Provided"
' 				If schl_grad_verif = "__" Then grad_date_verif = "Blank"
'
' 				If schl_fund = "1" Then school_funding = "1 - Not Attending in MN"
' 				If schl_fund = "2" Then school_funding = "2 - Attending Pub School"
' 				If schl_fund = "3" Then school_funding = "3 - Attending private/Parochial"
' 				If schl_fund = "4" Then school_funding = "4 - Not in Pre-12"
'
' 				If schl_elig = "01" Then school_elig_status = "01 - Under 18 or Over 50"
' 				If schl_elig = "02" Then school_elig_status = "02 - Disabled"
' 				If schl_elig = "03" Then school_elig_status = "03 - Not Higher Ed or < Halftime"
' 				If schl_elig = "04" Then school_elig_status = "04 - Employed 20 hrs/wk"
' 				If schl_elig = "05" Then school_elig_status = "05 - Work Study Program"
' 				If schl_elig = "06" Then school_elig_status = "06 - Dependant under 6"
' 				If schl_elig = "07" Then school_elig_status = "07 - Dep 6-11 No Child Care"
' 				If schl_elig = "09" Then school_elig_status = "09 - WIA, TAA, TRA or FSET"
' 				If schl_elig = "10" Then school_elig_status = "10 - Single Parent w/ Child < 12"
' 				If schl_elig = "99" Then school_elig_status = "99 - Not Eligible"
'
' 				If schl_higher_ed_yn = "Y" Then higher_ed = "Yes"
' 				If schl_higher_ed_yn = "N" Then higher_ed = "No"
' 				If schl_higher_ed_yn = "_" Then higher_ed = "Blank"
'
' 			End If
'
' 			Call navigate_to_MAXIS_screen("STAT", "STIN")		'===============================================================================================
' 			EMWriteScreen ref_number, 20, 76
' 			transmit
'
' 			EMreadScreen stin_version, 1, 2, 73
' 			If stin_version = "0" Then stin_exists = FALSE
' 			If stin_version = "1" Then stin_exists = TRUE
'
' 			If stin_exists = TRUE Then
' 				total_stin = 0
'
' 				stin_type_array = ARRAY("")
' 				stin_amount_array = ARRAY("")
' 				stin_avail_date_array = ARRAY("")
' 				stin_months_cov_array = ARRAY("")
' 				stin_verif_array = ARRAY("")
'
' 				stin_row = 8
' 				stin_counter = 0
' 				Do
' 					EMReadScreen stin_type, 2, stin_row, 27
' 					EMReadScreen stin_amount, 8, stin_row, 34
' 					EMReadScreen stin_date, 8, stin_row, 46
' 					EMReadScreen stin_month_one, 5, stin_row, 58
' 					EmReadscreen stin_month_two, 5, stin_row, 67
' 					EMReadScreen stin_verif, 1, stin_row, 76
'
'
' 					ReDim Preserve stin_type_array(stin_counter)
' 					ReDim Preserve stin_amount_array(stin_counter)
' 					ReDim Preserve stin_avail_date_array(stin_counter)
' 					ReDim Preserve stin_months_cov_array(stin_counter)
' 					ReDim Preserve stin_verif_array(stin_counter)
'
' 					If stin_type = "01" Then stin_type_array(stin_counter) = stin_type & " - Perkins Loan"
' 					If stin_type = "02" Then stin_type_array(stin_counter) = stin_type & " - Stafford Loan"
' 					If stin_type = "03" Then stin_type_array(stin_counter) = stin_type & " - Pell Grant"
' 					If stin_type = "04" Then stin_type_array(stin_counter) = stin_type & " - BIA Grant"
' 					If stin_type = "05" Then stin_type_array(stin_counter) = stin_type & " - SEOG"
' 					If stin_type = "06" Then stin_type_array(stin_counter) = stin_type & " - MN State Scholarship"
' 					If stin_type = "07" Then stin_type_array(stin_counter) = stin_type & " - Robert C Byrd Scholarship"
' 					If stin_type = "46" Then stin_type_array(stin_counter) = stin_type & " - Plus Loan (Deferred)"
' 					If stin_type = "16" Then stin_type_array(stin_counter) = stin_type & " - Plus Loan (Non-Deferred)"
' 					If stin_type = "47" Then stin_type_array(stin_counter) = stin_type & " - SLS (ALAS) Loan (Deferred)"
' 					If stin_type = "17" Then stin_type_array(stin_counter) = stin_type & " - SLS (ALAS) Loan (Non-Deferred)"
' 					If stin_type = "08" Then stin_type_array(stin_counter) = stin_type & " - Other Title IV Deferred Income"
' 					If stin_type = "09" Then stin_type_array(stin_counter) = stin_type & " - Other Title IV Grant"
' 					If stin_type = "10" Then stin_type_array(stin_counter) = stin_type & " - Other Title IV Scholarship"
' 					If stin_type = "11" Then stin_type_array(stin_counter) = stin_type & " - VA/GI Bill"
' 					If stin_type = "51" Then stin_type_array(stin_counter) = stin_type & " - VA/GI Bill (Earmarked)"
' 					If stin_type = "12" Then stin_type_array(stin_counter) = stin_type & " - Other Deferred Loan"
' 					If stin_type = "52" Then stin_type_array(stin_counter) = stin_type & " - Other Deferred Loan (Earmarked)"
' 					If stin_type = "13" Then stin_type_array(stin_counter) = stin_type & " - Other Grant"
' 					If stin_type = "53" Then stin_type_array(stin_counter) = stin_type & " - Other Grant (Earmarked)"
' 					If stin_type = "14" Then stin_type_array(stin_counter) = stin_type & " - Other Scholarship"
' 					If stin_type = "54" Then stin_type_array(stin_counter) = stin_type & " - Other Scholarship (Earmarked)"
' 					If stin_type = "15" Then stin_type_array(stin_counter) = stin_type & " - Other Aid"
' 					If stin_type = "55" Then stin_type_array(stin_counter) = stin_type & " - Other Aid (Earmarked)"
' 					If stin_type = "60" Then stin_type_array(stin_counter) = stin_type & " - MFIP Empl Svc (Earmarked)"
' 					If stin_type = "61" Then stin_type_array(stin_counter) = stin_type & " - WIOA, Unearned (Earmarked)"
' 					If stin_type = "18" Then stin_type_array(stin_counter) = stin_type & " - Other Exempt Loan"
' 					If stin_type = "62" Then stin_type_array(stin_counter) = stin_type & " - Tribal DSARLP"
'
' 					stin_amount_array(stin_counter) = trim(stin_amount)
'
' 					stin_avail_date_array(stin_counter) = replace(stin_date, " ", "/")
'
' 					stin_month_one = replace(stin_month_one, " ", "/")
' 					stin_month_two = replace(stin_month_two, " ", "/")
' 					stin_months_cov_array(stin_counter) = stin_month_one & " - " & stin_month_two
'
' 					If stin_verif = "1" Then stin_verif_array(stin_counter) = stin_verif & " - Award Letter"
' 					If stin_verif = "2" Then stin_verif_array(stin_counter) = stin_verif & " - DHS Financial Aid Form"
' 					If stin_verif = "3" Then stin_verif_array(stin_counter) = stin_verif & " - Student Profile Bulletin"
' 					If stin_verif = "4" Then stin_verif_array(stin_counter) = stin_verif & " - Pay Stubs"
' 					If stin_verif = "5" Then stin_verif_array(stin_counter) = stin_verif & " - Source Document"
' 					If stin_verif = "6" Then stin_verif_array(stin_counter) = stin_verif & " - Pend Out State Verif"
' 					If stin_verif = "7" Then stin_verif_array(stin_counter) = stin_verif & " - Other Document"
' 					If stin_verif = "N" Then stin_verif_array(stin_counter) = stin_verif & " - No Ver Prvd"
'
' 					stin_amount = stin_amount * 1
' 					total_stin = total_stin + stin_amount
'
' 					stin_row = stin_row + 1
' 					stin_counter = stin_counter + 1
'
' 					If stin_row = 18 Then
' 						PF20
' 						EMReadscreen last_page, 9, 24, 14
' 						If last_page = "LAST PAGE" Then Exit Do
' 						stin_row = 8
' 					End If
' 					EMReadScreen next_stin_type, 2, stin_row, 27
' 				Loop until next_stin_type = "__"
'
' 			End If
'
' 			Call navigate_to_MAXIS_screen("STAT", "STEC")		'===============================================================================================
' 			EMWriteScreen ref_number, 20, 76
' 			transmit
'
' 			EMreadScreen stec_version, 1, 2, 73
' 			If stec_version = "0" Then stec_exists = FALSE
' 			If stec_version = "1" Then stec_exists = TRUE
'
' 			If stec_exists = TRUE Then
' 				total_stec = 0
'
' 				stec_type_array = ARRAY("")
' 				stec_amount_array = ARRAY("")
' 				stec_months_cov_array = ARRAY("")
' 				stec_verif_array = ARRAY("")
' 				stec_earmarked_amount_array = ARRAY("")
' 				stec_earmarked_months_cov_array = ARRAY("")
'
' 				stec_row = 8
' 				stec_counter = 0
' 				Do
' 					EMReadScreen stec_type, 2, stec_row, 25
' 					EMReadScreen stec_amount, 8, stec_row, 31
' 					EMReadScreen stec_month_one, 5, stec_row, 41
' 					EMReadScreen stec_month_two, 5, stec_row, 48
' 					EMReadScreen stec_verif, 1, stec_row, 55
' 					EMReadScreen stec_earmarked_amount, 8, stec_row, 59
' 					EMReadScreen stec_earmarked_month_one, 2, stec_row, 69
' 					EMReadScreen stec_earmarked_month_two, 2, stec_row, 76
'
' 					ReDim Preserve stec_type_array(stec_counter)
' 					ReDim Preserve stec_amount_array(stec_counter)
' 					ReDim Preserve stec_months_cov_array(stec_counter)
' 					ReDim Preserve stec_verif_array(stec_counter)
' 					ReDim Preserve stec_earmarked_amount_array(stec_counter)
' 					ReDim Preserve stec_earmarked_months_cov_array(stec_counter)
'
' 					If stec_type = "" Then stec_type_array(stec_counter) = stec_type & " - "
'
' 					stec_amount_array(stec_counter) = trim(stec_amount)
'
' 					stec_month_one = replace(stec_month_one, " ", "/")
' 					stec_month_two = replace(stec_month_two, " ", "/")
' 					stec_months_cov_array(stec_counter) = stec_month_one & " - " & stec_month_two
'
' 					If stec_verif = "" Then stec_verif_array(stec_counter) = stec_verif & " - "
'
' 					stec_earmarked_amount_array(stec_counter) = trim(stec_earmarked_amount)
'
' 					stec_earmarked_month_one = replace(stec_earmarked_month_one, " ", "/")
' 					stec_earmarked_month_two = replace(stec_earmarked_month_two, " ", "/")
' 					stec_earmarked_months_cov_array(stec_counter) = stec_earmarked_month_one & " - " & stec_earmarked_month_two
'
' 					stec_amount = stec_amount * 1
' 					total_stec = total_stec + stec_amount
'
' 					stec_row = stec_row + 1
' 					stec_counter = stec_counter + 1
'
' 					If stec_row = 17 Then
' 						PF20
' 						EMReadscreen last_page, 9, 24, 14
' 						If last_page = "LAST PAGE" Then Exit Do
' 						stec_row = 8
' 					End If
' 					EMReadScreen next_stec_type, 2, stec_row, 25
' 				Loop until next_stec_type = "__"
' 			End If
'
' 			Call navigate_to_MAXIS_screen("STAT", "SHEL")		'===============================================================================================
' 			EMWriteScreen ref_number, 20, 76
' 			transmit
'
' 			EMreadScreen shel_version, 1, 2, 73
' 			If shel_version = "0" Then shel_exists = FALSE
' 			If shel_version = "1" Then shel_exists = TRUE
'
' 			If shel_exists = TRUE Then
' 				EMReadScreen shel_hud_subsidy_yn, 1, 6, 46
' 				EMReadScreen shel_shared_yn, 1, 6, 64
'
' 				EMReadScreen shel_paid_to, 25, 7, 50
'
' 				EMReadScreen shel_retro_rent_amount, 8, 11, 37
' 				EMReadScreen shel_retro_rent_verif, 2, 11, 48
' 				EMReadScreen shel_retro_lot_rent_amount, 8, 12, 37
' 				EMReadScreen shel_retro_lot_rent_verif, 2, 12, 48
' 				EMReadScreen shel_retro_mortgage_amount, 8, 13, 37
' 				EMReadScreen shel_retro_mortgage_verif, 2, 13, 48
' 				EMReadScreen shel_retro_insurance_amount, 8, 14, 37
' 				EMReadScreen shel_retro_insurance_verif, 2, 14, 48
' 				EMReadScreen shel_retro_taxes_amount, 8, 15, 37
' 				EMReadScreen shel_retro_taxes_verif, 2, 15, 48
' 				EMReadScreen shel_retro_room_amount, 8, 16, 37
' 				EMReadScreen shel_retro_room_verif, 2, 16, 48
' 				EMReadScreen shel_retro_garage_amount, 8, 17, 37
' 				EMReadScreen shel_retro_garage_verif, 2, 17, 48
' 				EMReadScreen shel_retro_subsidy_amount, 8, 18, 37
' 				EMReadScreen shel_retro_subsidy_verif, 2, 18, 48
'
' 				EMReadScreen shel_prosp_rent_amount, 8, 11, 56
' 				EMReadScreen shel_prosp_rent_verif, 2, 11, 67
' 				EMReadScreen shel_prosp_lot_rent_amount, 8, 12, 56
' 				EMReadScreen shel_prosp_lot_rent_verif, 2, 12, 67
' 				EMReadScreen shel_prosp_mortgage_amount, 8, 13, 56
' 				EMReadScreen shel_prosp_mortgage_verif, 2, 13, 67
' 				EMReadScreen shel_prosp_insurance_amount, 8, 14, 56
' 				EMReadScreen shel_prosp_insurance_verif, 2, 14, 67
' 				EMReadScreen shel_prosp_taxes_amount, 8, 15, 56
' 				EMReadScreen shel_prosp_taxes_verif, 2, 15, 67
' 				EMReadScreen shel_prosp_room_amount, 8, 16, 56
' 				EMReadScreen shel_prosp_room_verif, 2, 16, 67
' 				EMReadScreen shel_prosp_garage_amount, 8, 17, 56
' 				EMReadScreen shel_prosp_garage_verif, 2, 17, 67
' 				EMReadScreen shel_prosp_subsidy_amount, 8, 18, 56
' 				EMReadScreen shel_prosp_subsidy_verif, 2, 18, 67
'
' 				shel_paid_to = replace(shel_paid_to, "_", "")
'
' 				shel_retro_rent_amount = trim(replace(shel_retro_rent_amount, "_", ""))
' 				shel_retro_lot_rent_amount = trim(replace(shel_retro_lot_rent_amount, "_", ""))
' 				shel_retro_mortgage_amount = trim(replace(shel_retro_mortgage_amount, "_", ""))
' 				shel_retro_insurance_amount = trim(replace(shel_retro_insurance_amount, "_", ""))
' 				shel_retro_taxes_amount = trim(replace(shel_retro_taxes_amount, "_", ""))
' 				shel_retro_room_amount = trim(replace(shel_retro_room_amount, "_", ""))
' 				shel_retro_garage_amount = trim(replace(shel_retro_garage_amount, "_", ""))
' 				shel_retro_subsidy_amount = trim(replace(shel_retro_subsidy_amount, "_", ""))
'
' 				shel_prosp_rent_amount = trim(replace(shel_prosp_rent_amount, "_", ""))
' 				shel_prosp_lot_rent_amount = trim(replace(shel_prosp_lot_rent_amount, "_", ""))
' 				shel_prosp_mortgage_amount = trim(replace(shel_prosp_mortgage_amount, "_", ""))
' 				shel_prosp_insurance_amount = trim(replace(shel_prosp_insurance_amount, "_", ""))
' 				shel_prosp_taxes_amount = trim(replace(shel_prosp_taxes_amount, "_", ""))
' 				shel_prosp_room_amount = trim(replace(shel_prosp_room_amount, "_", ""))
' 				shel_prosp_garage_amount = trim(replace(shel_prosp_garage_amount, "_", ""))
' 				shel_prosp_subsidy_amount = trim(replace(shel_prosp_subsidy_amount, "_", ""))
'
' 				If shel_prosp_rent_amount <> "" Then shel_summary = shel_summary & " Rent: $" & shel_prosp_rent_amount & " - Verif: " & shel_prosp_rent_verif & " | "
' 				If shel_prosp_lot_rent_amount <> "" Then shel_summary = shel_summary & " Lot Rent: $" & shel_prosp_lot_rent_amount & " - Verif: " & shel_prosp_lot_rent_verif & " | "
' 				If shel_prosp_mortgage_amount <> "" Then shel_summary = shel_summary & " Mortgage: $" & shel_prosp_mortgage_amount & " - Verif: " & shel_prosp_mortgage_verif & " | "
' 				If shel_prosp_insurance_amount <> "" Then shel_summary = shel_summary & " Insurance: $" & shel_prosp_insurance_amount & " - Verif: " & shel_prosp_insurance_verif & " | "
' 				If shel_prosp_taxes_amount <> "" Then shel_summary = shel_summary & " Taxes: $" & shel_prosp_taxes_amount & " - Verif: " & shel_prosp_taxes_verif & " | "
' 				If shel_prosp_room_amount <> "" Then shel_summary = shel_summary & " Room: $" & shel_prosp_room_amount & " - Verif: " & shel_prosp_room_verif & " | "
' 				If shel_prosp_garage_amount <> "" Then shel_summary = shel_summary & " Garage: $" & shel_prosp_garage_amount & " - Verif: " & shel_prosp_garage_verif & " | "
' 				If shel_prosp_subsidy_amount <> "" Then shel_summary = shel_summary & " Subsidy: $" & shel_prosp_subsidy_amount & " - Verif: " & shel_prosp_subsidy_verif & " | "
'
' 				If shel_retro_rent_verif = "SF" Then shel_retro_rent_verif = shel_retro_rent_verif & " - Shelter Form"
' 				If shel_retro_rent_verif = "LE" Then shel_retro_rent_verif = shel_retro_rent_verif & " - Lease"
' 				If shel_retro_rent_verif = "RE" Then shel_retro_rent_verif = shel_retro_rent_verif & " - Rent Receipts"
' 				If shel_retro_rent_verif = "OT" Then shel_retro_rent_verif = shel_retro_rent_verif & " - Other Document"
' 				If shel_retro_rent_verif = "NC" Then shel_retro_rent_verif = shel_retro_rent_verif & " - Not Verif, Neg Impact"
' 				If shel_retro_rent_verif = "PC" Then shel_retro_rent_verif = shel_retro_rent_verif & " - Not Verif, Pos Impact"
' 				If shel_retro_rent_verif = "NO" Then shel_retro_rent_verif = shel_retro_rent_verif & " - No Verif Provided"
'
' 				If shel_retro_lot_rent_verif = "LE" Then shel_retro_lot_rent_verif = shel_retro_lot_rent_verif & " - Lease"
' 				If shel_retro_lot_rent_verif = "RE" Then shel_retro_lot_rent_verif = shel_retro_lot_rent_verif & " - Rent Receipts"
' 				If shel_retro_lot_rent_verif = "BI" Then shel_retro_lot_rent_verif = shel_retro_lot_rent_verif & " - Billing Statement"
' 				If shel_retro_lot_rent_verif = "OT" Then shel_retro_lot_rent_verif = shel_retro_lot_rent_verif & " - Other Document"
' 				If shel_retro_lot_rent_verif = "NC" Then shel_retro_lot_rent_verif = shel_retro_lot_rent_verif & " - Not Verif, Neg Impact"
' 				If shel_retro_lot_rent_verif = "PC" Then shel_retro_lot_rent_verif = shel_retro_lot_rent_verif & " - Not Verif, Pos Impact"
' 				If shel_retro_lot_rent_verif = "NO" Then shel_retro_lot_rent_verif = shel_retro_lot_rent_verif & " - No Verif Provided"
'
' 				If shel_retro_mortgage_verif = "MO" Then shel_retro_mortgage_verif = shel_retro_mortgage_verif & " - Mortgage Payment"
' 				If shel_retro_mortgage_verif = "CD" Then shel_retro_mortgage_verif = shel_retro_mortgage_verif & " - Contract for Deed"
' 				If shel_retro_mortgage_verif = "OT" Then shel_retro_mortgage_verif = shel_retro_mortgage_verif & " - Other Document"
' 				If shel_retro_mortgage_verif = "NC" Then shel_retro_mortgage_verif = shel_retro_mortgage_verif & " - Not Verif, Neg Impact"
' 				If shel_retro_mortgage_verif = "PC" Then shel_retro_mortgage_verif = shel_retro_mortgage_verif & " - Not Verif, Pos Impact"
' 				If shel_retro_mortgage_verif = "NO" Then shel_retro_mortgage_verif = shel_retro_mortgage_verif & " - No Verif Provided"
'
' 				If shel_retro_insurance_verif = "BI" Then shel_retro_insurance_verif = shel_retro_insurance_verif & " - Billing Statement"
' 				If shel_retro_insurance_verif = "OT" Then shel_retro_insurance_verif = shel_retro_insurance_verif & " - Other Document"
' 				If shel_retro_insurance_verif = "NC" Then shel_retro_insurance_verif = shel_retro_insurance_verif & " - Not Verif, Neg Impact"
' 				If shel_retro_insurance_verif = "PC" Then shel_retro_insurance_verif = shel_retro_insurance_verif & " - Not Verif, Pos Impact"
' 				If shel_retro_insurance_verif = "NO" Then shel_retro_insurance_verif = shel_retro_insurance_verif & " - No Verif Provided"
'
' 				If shel_retro_taxes_verif = "TX" Then shel_retro_taxes_verif = shel_retro_taxes_verif & " - Property Tax Statement"
' 				If shel_retro_taxes_verif = "OT" Then shel_retro_taxes_verif = shel_retro_taxes_verif & " - Other Document"
' 				If shel_retro_taxes_verif = "NC" Then shel_retro_taxes_verif = shel_retro_taxes_verif & " - Not Verif, Neg Impact"
' 				If shel_retro_taxes_verif = "PC" Then shel_retro_taxes_verif = shel_retro_taxes_verif & " - Not Verif, Pos Impact"
' 				If shel_retro_taxes_verif = "NO" Then shel_retro_taxes_verif = shel_retro_taxes_verif & " - No Verif Provided"
'
' 				If shel_retro_room_verif = "SF" Then shel_retro_room_verif = shel_retro_room_verif & " - Shelter Form"
' 				If shel_retro_room_verif = "LE" Then shel_retro_room_verif = shel_retro_room_verif & " - Lease"
' 				If shel_retro_room_verif = "RE" Then shel_retro_room_verif = shel_retro_room_verif & " - Rent Receipts"
' 				If shel_retro_room_verif = "OT" Then shel_retro_room_verif = shel_retro_room_verif & " - Other Document"
' 				If shel_retro_room_verif = "NC" Then shel_retro_room_verif = shel_retro_room_verif & " - Not Verif, Neg Impact"
' 				If shel_retro_room_verif = "PC" Then shel_retro_room_verif = shel_retro_room_verif & " - Not Verif, Pos Impact"
' 				If shel_retro_room_verif = "NO" Then shel_retro_room_verif = shel_retro_room_verif & " - No Verif Provided"
'
' 				If shel_retro_garage_verif = "SF" Then shel_retro_garage_verif = shel_retro_garage_verif & " - Shelter Form"
' 				If shel_retro_garage_verif = "LE" Then shel_retro_garage_verif = shel_retro_garage_verif & " - Lease"
' 				If shel_retro_garage_verif = "RE" Then shel_retro_garage_verif = shel_retro_garage_verif & " - Rent Receipts"
' 				If shel_retro_garage_verif = "OT" Then shel_retro_garage_verif = shel_retro_garage_verif & " - Other Document"
' 				If shel_retro_garage_verif = "NC" Then shel_retro_garage_verif = shel_retro_garage_verif & " - Not Verif, Neg Impact"
' 				If shel_retro_garage_verif = "PC" Then shel_retro_garage_verif = shel_retro_garage_verif & " - Not Verif, Pos Impact"
' 				If shel_retro_garage_verif = "NO" Then shel_retro_garage_verif = shel_retro_garage_verif & " - No Verif Provided"
'
' 				If shel_retro_subsidy_verif = "SF" Then shel_retro_subsidy_verif = shel_retro_subsidy_verif & " - Shelter Form"
' 				If shel_retro_subsidy_verif = "LE" Then shel_retro_subsidy_verif = shel_retro_subsidy_verif & " - Lease"
' 				If shel_retro_subsidy_verif = "OT" Then shel_retro_subsidy_verif = shel_retro_subsidy_verif & " - Other Document"
' 				If shel_retro_subsidy_verif = "NO" Then shel_retro_subsidy_verif = shel_retro_subsidy_verif & " - No Verif Provided"
'
'
' 				If shel_prosp_rent_verif = "SF" Then shel_prosp_rent_verif = shel_prosp_rent_verif & " - Shelter Form"
' 				If shel_prosp_rent_verif = "LE" Then shel_prosp_rent_verif = shel_prosp_rent_verif & " - Lease"
' 				If shel_prosp_rent_verif = "RE" Then shel_prosp_rent_verif = shel_prosp_rent_verif & " - Rent Receipts"
' 				If shel_prosp_rent_verif = "OT" Then shel_prosp_rent_verif = shel_prosp_rent_verif & " - Other Document"
' 				If shel_prosp_rent_verif = "NC" Then shel_prosp_rent_verif = shel_prosp_rent_verif & " - Not Verif, Neg Impact"
' 				If shel_prosp_rent_verif = "PC" Then shel_prosp_rent_verif = shel_prosp_rent_verif & " - Not Verif, Pos Impact"
' 				If shel_prosp_rent_verif = "NO" Then shel_prosp_rent_verif = shel_prosp_rent_verif & " - No Verif Provided"
'
' 				If shel_prosp_lot_rent_verif = "LE" Then shel_prosp_lot_rent_verif = shel_prosp_lot_rent_verif & " - Lease"
' 				If shel_prosp_lot_rent_verif = "RE" Then shel_prosp_lot_rent_verif = shel_prosp_lot_rent_verif & " - Rent Receipts"
' 				If shel_prosp_lot_rent_verif = "BI" Then shel_prosp_lot_rent_verif = shel_prosp_lot_rent_verif & " - Billing Statement"
' 				If shel_prosp_lot_rent_verif = "OT" Then shel_prosp_lot_rent_verif = shel_prosp_lot_rent_verif & " - Other Document"
' 				If shel_prosp_lot_rent_verif = "NC" Then shel_prosp_lot_rent_verif = shel_prosp_lot_rent_verif & " - Not Verif, Neg Impact"
' 				If shel_prosp_lot_rent_verif = "PC" Then shel_prosp_lot_rent_verif = shel_prosp_lot_rent_verif & " - Not Verif, Pos Impact"
' 				If shel_prosp_lot_rent_verif = "NO" Then shel_prosp_lot_rent_verif = shel_prosp_lot_rent_verif & " - No Verif Provided"
'
' 				If shel_prosp_mortgage_verif = "MO" Then shel_prosp_mortgage_verif = shel_prosp_mortgage_verif & " - Mortgage Payment"
' 				If shel_prosp_mortgage_verif = "CD" Then shel_prosp_mortgage_verif = shel_prosp_mortgage_verif & " - Contract for Deed"
' 				If shel_prosp_mortgage_verif = "OT" Then shel_prosp_mortgage_verif = shel_prosp_mortgage_verif & " - Other Document"
' 				If shel_prosp_mortgage_verif = "NC" Then shel_prosp_mortgage_verif = shel_prosp_mortgage_verif & " - Not Verif, Neg Impact"
' 				If shel_prosp_mortgage_verif = "PC" Then shel_prosp_mortgage_verif = shel_prosp_mortgage_verif & " - Not Verif, Pos Impact"
' 				If shel_prosp_mortgage_verif = "NO" Then shel_prosp_mortgage_verif = shel_prosp_mortgage_verif & " - No Verif Provided"
'
' 				If shel_prosp_insurance_verif = "BI" Then shel_prosp_insurance_verif = shel_prosp_insurance_verif & " - Billing Statement"
' 				If shel_prosp_insurance_verif = "OT" Then shel_prosp_insurance_verif = shel_prosp_insurance_verif & " - Other Document"
' 				If shel_prosp_insurance_verif = "NC" Then shel_prosp_insurance_verif = shel_prosp_insurance_verif & " - Not Verif, Neg Impact"
' 				If shel_prosp_insurance_verif = "PC" Then shel_prosp_insurance_verif = shel_prosp_insurance_verif & " - Not Verif, Pos Impact"
' 				If shel_prosp_insurance_verif = "NO" Then shel_prosp_insurance_verif = shel_prosp_insurance_verif & " - No Verif Provided"
'
' 				If shel_prosp_taxes_verif = "TX" Then shel_prosp_taxes_verif = shel_prosp_taxes_verif & " - Property Tax Statement"
' 				If shel_prosp_taxes_verif = "OT" Then shel_prosp_taxes_verif = shel_prosp_taxes_verif & " - Other Document"
' 				If shel_prosp_taxes_verif = "NC" Then shel_prosp_taxes_verif = shel_prosp_taxes_verif & " - Not Verif, Neg Impact"
' 				If shel_prosp_taxes_verif = "PC" Then shel_prosp_taxes_verif = shel_prosp_taxes_verif & " - Not Verif, Pos Impact"
' 				If shel_prosp_taxes_verif = "NO" Then shel_prosp_taxes_verif = shel_prosp_taxes_verif & " - No Verif Provided"
'
' 				If shel_prosp_room_verif = "SF" Then shel_prosp_room_verif = shel_prosp_room_verif & " - Shelter Form"
' 				If shel_prosp_room_verif = "LE" Then shel_prosp_room_verif = shel_prosp_room_verif & " - Lease"
' 				If shel_prosp_room_verif = "RE" Then shel_prosp_room_verif = shel_prosp_room_verif & " - Rent Receipts"
' 				If shel_prosp_room_verif = "OT" Then shel_prosp_room_verif = shel_prosp_room_verif & " - Other Document"
' 				If shel_prosp_room_verif = "NC" Then shel_prosp_room_verif = shel_prosp_room_verif & " - Not Verif, Neg Impact"
' 				If shel_prosp_room_verif = "PC" Then shel_prosp_room_verif = shel_prosp_room_verif & " - Not Verif, Pos Impact"
' 				If shel_prosp_room_verif = "NO" Then shel_prosp_room_verif = shel_prosp_room_verif & " - No Verif Provided"
'
' 				If shel_prosp_garage_verif = "SF" Then shel_prosp_garage_verif = shel_prosp_garage_verif & " - Shelter Form"
' 				If shel_prosp_garage_verif = "LE" Then shel_prosp_garage_verif = shel_prosp_garage_verif & " - Lease"
' 				If shel_prosp_garage_verif = "RE" Then shel_prosp_garage_verif = shel_prosp_garage_verif & " - Rent Receipts"
' 				If shel_prosp_garage_verif = "OT" Then shel_prosp_garage_verif = shel_prosp_garage_verif & " - Other Document"
' 				If shel_prosp_garage_verif = "NC" Then shel_prosp_garage_verif = shel_prosp_garage_verif & " - Not Verif, Neg Impact"
' 				If shel_prosp_garage_verif = "PC" Then shel_prosp_garage_verif = shel_prosp_garage_verif & " - Not Verif, Pos Impact"
' 				If shel_prosp_garage_verif = "NO" Then shel_prosp_garage_verif = shel_prosp_garage_verif & " - No Verif Provided"
'
' 				If shel_prosp_subsidy_verif = "SF" Then shel_prosp_subsidy_verif = shel_prosp_subsidy_verif & " - Shelter Form"
' 				If shel_prosp_subsidy_verif = "LE" Then shel_prosp_subsidy_verif = shel_prosp_subsidy_verif & " - Lease"
' 				If shel_prosp_subsidy_verif = "OT" Then shel_prosp_subsidy_verif = shel_prosp_subsidy_verif & " - Other Document"
' 				If shel_prosp_subsidy_verif = "NO" Then shel_prosp_subsidy_verif = shel_prosp_subsidy_verif & " - No Verif Provided"
'
' 			End If
'
' 			Call navigate_to_MAXIS_screen("STAT", "STWK")		'===============================================================================================
' 			EMWriteScreen ref_number, 20, 76
' 			transmit
'
' 			EMreadScreen stwk_version, 1, 2, 73
' 			If stwk_version = "0" Then stwk_exists = FALSE
' 			If stwk_version = "1" Then stwk_exists = TRUE
'
' 			If stwk_exists = TRUE Then
' 				EMReadScreen stwk_employer, 30, 6, 46
' 				EMReadScreen stwk_work_stop_date, 8, 7, 46
' 				EMReadScreen stwk_income_stop_date, 8, 8, 46
' 				EMReadScreen stwk_verification, 1, 7, 63
' 				EMReadScreen stwk_refused_employment, 1, 8, 78
' 				EMReadScreen stwk_vol_quit, 1, 10, 46
' 				EMReadScreen stwk_refused_employment_date, 8, 10, 72
' 				EMReadScreen stwk_cash_good_cause_yn, 1, 12, 52
' 				EMReadScreen stwk_grh_good_cause_yn, 1, 12, 60
' 				EMReadScreen stwk_snap_good_cause_yn, 1, 12, 67
' 				EMReadScreen stwk_snap_pwe, 1, 14, 46
' 				EMReadScreen stwk_ma_epd_extension, 1, 16, 46
'
' 				stwk_employer = replace(stwk_employer, "_", "")
' 				stwk_work_stop_date = replace(stwk_work_stop_date, " ", "/")
' 				stwk_income_stop_date = replace(stwk_income_stop_date, " ", "/")
' 				If stwk_verification = "1" Then stwk_verification = "Employers Statement"
' 				If stwk_verification = "2" Then stwk_verification = "Seperation Notice"
' 				If stwk_verification = "3" Then stwk_verification = "Colateral Statement"
' 				If stwk_verification = "4" Then stwk_verification = "Other Document"
' 				If stwk_verification = "N" Then stwk_verification = "No Verif Provided"
' 				If stwk_verification = "_" Then stwk_verification = "Blank"
' 				If stwk_verification = "?" Then stwk_verification = "Postponed Verif"
' 				stwk_refused_employment_date = replace(stwk_refused_employment_date, " ", "/")
' 				stwk_summary = "Work ended at " & stwk_employer & " on " & stwk_work_stop_date
'
' 			End If
'
' 			Call navigate_to_MAXIS_screen("STAT", "FMED")		'===============================================================================================
' 			EMWriteScreen ref_number, 20, 76
' 			transmit
'
' 			EMreadScreen fmed_version, 1, 2, 73
' 			If fmed_version = "0" Then fmed_exists = FALSE
' 			If fmed_version = "1" Then fmed_exists = TRUE
'
' 			If fmed_exists = TRUE Then
' 				EMReadScreen fmed_miles, 4, 17, 34
' 				EMReadScreen fmed_rate, 6, 17, 58
' 				EMReadScreen fmed_milage_expense, 8, 17, 70
'
' 				panel_row = 9
' 				fmed_count = 0
' 				scroll_page = 1
' 				Do
' 					EMReadScreen the_type, 2, panel_row, 25
'
' 					If the_type <> "__" Then
' 						' ReDim Preserve fmed_expense_array(fmed_count, fmed_notes)
' 						ReDim Preserve fmed_page(fmed_count)
' 						ReDim Preserve fmed_row(fmed_count)
' 						ReDim Preserve fmed_type(fmed_count)
' 						ReDim Preserve fmed_verif(fmed_count)
' 						ReDim Preserve fmed_ref(fmed_count)
' 						ReDim Preserve fmed_catgry(fmed_count)
' 						ReDim Preserve fmed_begin(fmed_count)
' 						ReDim Preserve fmed_end(fmed_count)
' 						ReDim Preserve fmed_expense(fmed_count)
' 						ReDim Preserve fmed_notes(fmed_count)
'
' 						EMReadScreen the_ver, 2, panel_row, 32
' 						EMReadScreen the_ref, 2, panel_row, 38
' 						EMReadScreen the_cat, 1, panel_row, 44
' 						EMReadScreen the_begin, 5, panel_row, 50
' 						EMReadScreen the_end, 5, panel_row, 60
' 						EMReadScreen the_amt, 8, panel_row, 70
'
' 						fmed_page(fmed_count) = scroll_page
' 						fmed_row(fmed_count) = panel_row
'
' 						If the_type = "01" Then fmed_type(fmed_count) = "01 Nursing Home"
' 						If the_type = "02" Then fmed_type(fmed_count) = "02 Hosp/Clinic"
' 						If the_type = "03" Then fmed_type(fmed_count) = "03 Physicians"
' 						If the_type = "04" Then fmed_type(fmed_count) = "04 Prescriptions"
' 						If the_type = "05" Then fmed_type(fmed_count) = "05 Ins Premiums"
' 						If the_type = "06" Then fmed_type(fmed_count) = "06 Dental"
' 						If the_type = "07" Then fmed_type(fmed_count) = "07 Medical Trans/Flat Amount"
' 						If the_type = "08" Then fmed_type(fmed_count) = "08 Vision Care"
' 						If the_type = "09" Then fmed_type(fmed_count) = "09 Medicare Prem"
' 						If the_type = "10" Then fmed_type(fmed_count) = "10 Mo Spdwn Amt/Waiver Oblig"
' 						If the_type = "11" Then fmed_type(fmed_count) = "11 Home Care"
' 						If the_type = "12" Then fmed_type(fmed_count) = "12 Medical Trans/Mileage Calc"
' 						If the_type = "15" Then fmed_type(fmed_count) = "15 Medi Part D Premium"
'
' 						If the_ver = "BI" Then fmed_verif(fmed_count) = "BI Billing Stmt"
' 						If the_ver = "EB" Then fmed_verif(fmed_count) = "EB Expl Of Bnft (Medicare/Ins)"
' 						If the_ver = "CL" Then fmed_verif(fmed_count) = "CL Client Stmt Med Trans Only"
' 						If the_ver = "OS" Then fmed_verif(fmed_count) = "OS Pend Out State Verification"
' 						If the_ver = "OT" Then fmed_verif(fmed_count) = "OT Other Document"
' 						If the_ver = "NO" Then fmed_verif(fmed_count) = "NO No Ver Prvd"
' 						If the_ver = "MX" Then fmed_verif(fmed_count) = "MX System Entered Ver By SSA"
'
' 						fmed_ref(fmed_count) = the_ref
'
' 						If the_cat = "1" Then fmed_catgry(fmed_count) = "1 HH Member"
' 						If the_cat = "2" Then fmed_catgry(fmed_count) = "2 Former Aged/Disa HH Mbr In NF Or Hospital"
' 						If the_cat = "3" Then fmed_catgry(fmed_count) = "3 Former Aged/Disa HH Decd"
' 						If the_cat = "4" Then fmed_catgry(fmed_count) = "4 Other Eligible"
'
' 						fmed_begin(fmed_count) = replace(the_begin, " ", "/")
' 						fmed_end(fmed_count) = replace(the_end, " ", "/")
' 						fmed_expense(fmed_count) = trim(the_amt)
'
' 						panel_row = panel_row + 1
' 						fmed_count = fmed_count + 1
' 						If panel_row = 15 Then
' 							pf20
' 							scroll_page = scroll_page + 1
' 							panel_row = 9
' 							EMReadScreen end_of_list, 9, 24, 14
' 							If end_of_list = "LAST PAGE" Then Exit Do
' 						End If
' 					End If
' 				Loop until panel_type = "__"
' 			End If
'
' 			Call navigate_to_MAXIS_screen("STAT", "PARE")		'===============================================================================================
' 			EMWriteScreen ref_number, 20, 76
' 			transmit
'
' 			EMreadScreen pare_version, 1, 2, 73
' 			If pare_version = "0" Then pare_exists = FALSE
' 			If pare_version = "1" Then pare_exists = TRUE
'
' 			If pare_exists = TRUE Then
' 				pare_row = 8
' 				pare_array_count = 0
'
' 				Do
' 					EMReadScreen panel_child_ref_number, 2, pare_row, 24
' 					EMReadScreen panel_child_name, 25, pare_row, 27
' 					EMReadScreen panel_rela_type, 1, pare_row, 53
' 					EMReadScreen panel_rela_verif, 2, pare_row, 71
'
' 					If panel_child_ref_number <> "__" Then
' 						ReDim preserve pare_child_ref_nbr(pare_array_count)
' 						ReDim preserve pare_child_name(pare_array_count)
' 						ReDim preserve pare_child_member_index(pare_array_count)
' 						ReDim preserve pare_relationship_type(pare_array_count)
' 						ReDim preserve pare_verification(pare_array_count)
'
' 						pare_child_ref_nbr(pare_array_count) = panel_child_ref_number
' 						pare_child_name(pare_array_count) = trim(panel_child_name)
'
' 						' pare_child_member_index(pare_array_count)
'
' 						If panel_rela_type = "1" Then pare_relationship_type(pare_array_count) = "1 - Birth/Adopted Parent"
' 						If panel_rela_type = "2" Then pare_relationship_type(pare_array_count) = "2 - Stepchild"
' 						If panel_rela_type = "3" Then pare_relationship_type(pare_array_count) = "3 - Grandchild"
' 						If panel_rela_type = "4" Then pare_relationship_type(pare_array_count) = "4 - Relative Caregiver"
' 						If panel_rela_type = "5" Then pare_relationship_type(pare_array_count) = "5 - Foster Child"
' 						If panel_rela_type = "6" Then pare_relationship_type(pare_array_count) = "6 - Non-related Caregiver"
' 						If panel_rela_type = "7" Then pare_relationship_type(pare_array_count) = "7 - Legal Guardian"
' 						If panel_rela_type = "8" Then pare_relationship_type(pare_array_count) = "8 - Other Relative"
'
' 						If panel_rela_verif = "BC" Then pare_verification(pare_array_count) = "BC - Birth Certificate"
' 						If panel_rela_verif = "AR" Then pare_verification(pare_array_count) = "AR - Adoption Records"
' 						If panel_rela_verif = "LG" Then pare_verification(pare_array_count) = "LG - Legal Guardian"
' 						If panel_rela_verif = "RE" Then pare_verification(pare_array_count) = "RE - Religious Records"
' 						If panel_rela_verif = "HR" Then pare_verification(pare_array_count) = "HR - Hospital Records"
' 						If panel_rela_verif = "RP" Then pare_verification(pare_array_count) = "RP - Recognition of Parentage"
' 						If panel_rela_verif = "OT" Then pare_verification(pare_array_count) = "OT - Other Verification"
' 						If panel_rela_verif = "NO" Then pare_verification(pare_array_count) = "NO - No Verif Provided"
' 						If panel_rela_verif = "__" Then pare_verification(pare_array_count) = "Blank"
' 						If panel_rela_verif = "?_" Then pare_verification(pare_array_count) = "Delayed Verification"
' 					End If
'
' 					pare_row = pare_row + 1
' 					pare_array_count = pare_array_count + 1
' 					If pare_row = 18 Then
' 						pare_row = 8
' 						PF20
' 						EMReadScreen end_of_list, 9, 24, 14
' 						If end_of_list = "LAST PAGE" then Exit Do
' 					End If
' 				Loop until panel_child_ref_number = "__"
' 			End If
'
' 			Call navigate_to_MAXIS_screen("STAT", "PDED")		'===============================================================================================
' 			EMWriteScreen ref_number, 20, 76
' 			transmit
'
' 			EMreadScreen pded_version, 1, 2, 73
' 			If pded_version = "0" Then pded_exists = FALSE
' 			If pded_version = "1" Then pded_exists = TRUE
'
' 			If pded_exists = TRUE Then
' 				EMReadScreen pded_guardian_fee, 8, 15, 44
' 				EMReadScreen pded_rep_payee_fee, 8, 15, 70
' 				EMReadScreen pded_shel_spec_need, 1, 18, 78
'
' 				pded_guardian_fee = replace(pded_guardian_fee, "_", "")
' 				pded_guardian_fee = trim(pded_guardian_fee)
' 				' MsgBox pded_rep_payee_fee & " 1"
' 				pded_rep_payee_fee = replace(pded_rep_payee_fee, "_", "")
' 				pded_rep_payee_fee = trim(pded_rep_payee_fee)
' 				' MsgBox pded_rep_payee_fee & " 2"
'
' 				If pded_shel_spec_need = "Y" Then pded_shel_spec_need = "Yes"
' 				If pded_shel_spec_need = "N" Then pded_shel_spec_need = "No"
' 				If pded_shel_spec_need = "_" Then pded_shel_spec_need = ""
' 			End If
'
'
' 			Call navigate_to_MAXIS_screen("STAT", "DIET")		'===============================================================================================
' 			EMWriteScreen ref_number, 20, 76
' 			transmit
'
' 			EMreadScreen diet_version, 1, 2, 73
' 			If diet_version = "0" Then diet_exists = FALSE
' 			If diet_version = "1" Then diet_exists = TRUE
'
' 			If diet_exists = TRUE Then
' 				EMReadScreen diet_mf_type_one, 2, 8, 40
' 				EMReadScreen diet_mf_verif_one, 1, 8, 51
' 				EMReadScreen diet_mf_type_two, 2, 9, 40
' 				EMReadScreen diet_mf_verif_two, 1, 9, 51
'
' 				EMReadScreen diet_msa_type_one, 2, 11, 40
' 				EMReadScreen diet_msa_verif_one, 1, 11, 51
' 				EMReadScreen diet_msa_type_two, 2, 12, 40
' 				EMReadScreen diet_msa_verif_two, 1, 12, 51
' 				EMReadScreen diet_msa_type_three, 2, 13, 40
' 				EMReadScreen diet_msa_verif_three, 1, 13, 51
' 				EMReadScreen diet_msa_type_four, 2, 14, 40
' 				EMReadScreen diet_msa_verif_four, 1, 14, 51
' 				EMReadScreen diet_msa_type_five, 2, 15, 40
' 				EMReadScreen diet_msa_verif_five, 1, 15, 51
' 				EMReadScreen diet_msa_type_six, 2, 16, 40
' 				EMReadScreen diet_msa_verif_six, 1, 16, 51
' 				EMReadScreen diet_msa_type_seven, 2, 17, 40
' 				EMReadScreen diet_msa_verif_seven, 1, 17, 51
' 				EMReadScreen diet_msa_type_eight, 2, 18, 40
' 				EMReadScreen diet_msa_verif_eight, 1, 18, 51
'
' 				If diet_mf_type_one = "01" Then diet_mf_type_one = "01 - High Protein > 79 grams/day"
' 				If diet_mf_type_one = "02" Then diet_mf_type_one = "02 - Control Protein 40-60 grams/day"
' 				If diet_mf_type_one = "03" Then diet_mf_type_one = "03 - Control Protein < 40 grams/day"
' 				If diet_mf_type_one = "04" Then diet_mf_type_one = "04 - Lo Cholesterol"
' 				If diet_mf_type_one = "05" Then diet_mf_type_one = "05 - High Residue"
' 				If diet_mf_type_one = "06" Then diet_mf_type_one = "06 - Pregnancy and Lactation"
' 				If diet_mf_type_one = "07" Then diet_mf_type_one = "07 - Gluten Free"
' 				If diet_mf_type_one = "08" Then diet_mf_type_one = "08 - Lactose Free"
' 				If diet_mf_type_one = "09" Then diet_mf_type_one = "09 - Anti-Dumping"
' 				If diet_mf_type_one = "10" Then diet_mf_type_one = "10 - Hypoglycemic"
' 				If diet_mf_type_one = "11" Then diet_mf_type_one = "11 - Ketogenic"
' 				If diet_mf_type_one = "__" Then diet_mf_type_one = ""
'
' 				If diet_mf_type_two = "01" Then diet_mf_type_two = "01 - High Protein > 79 grams/day"
' 				If diet_mf_type_two = "02" Then diet_mf_type_two = "02 - Control Protein 40-60 grams/day"
' 				If diet_mf_type_two = "03" Then diet_mf_type_two = "03 - Control Protein < 40 grams/day"
' 				If diet_mf_type_two = "04" Then diet_mf_type_two = "04 - Lo Cholesterol"
' 				If diet_mf_type_two = "05" Then diet_mf_type_two = "05 - High Residue"
' 				If diet_mf_type_two = "06" Then diet_mf_type_two = "06 - Pregnancy and Lactation"
' 				If diet_mf_type_two = "07" Then diet_mf_type_two = "07 - Gluten Free"
' 				If diet_mf_type_two = "08" Then diet_mf_type_two = "08 - Lactose Free"
' 				If diet_mf_type_two = "09" Then diet_mf_type_two = "09 - Anti-Dumping"
' 				If diet_mf_type_two = "10" Then diet_mf_type_two = "10 - Hypoglycemic"
' 				If diet_mf_type_two = "11" Then diet_mf_type_two = "11 - Ketogenic"
' 				If diet_mf_type_two = "__" Then diet_mf_type_two = ""
'
'
' 				If diet_msa_type_one = "01" Then diet_msa_type_one = "01 - High Protein > 79 grams/day"
' 				If diet_msa_type_one = "02" Then diet_msa_type_one = "02 - Control Protein 40-60 grams/day"
' 				If diet_msa_type_one = "03" Then diet_msa_type_one = "03 - Control Protein < 40 grams/day"
' 				If diet_msa_type_one = "04" Then diet_msa_type_one = "04 - Lo Cholesterol"
' 				If diet_msa_type_one = "05" Then diet_msa_type_one = "05 - High Residue"
' 				If diet_msa_type_one = "06" Then diet_msa_type_one = "06 - Pregnancy and Lactation"
' 				If diet_msa_type_one = "07" Then diet_msa_type_one = "07 - Gluten Free"
' 				If diet_msa_type_one = "08" Then diet_msa_type_one = "08 - Lactose Free"
' 				If diet_msa_type_one = "09" Then diet_msa_type_one = "09 - Anti-Dumping"
' 				If diet_msa_type_one = "10" Then diet_msa_type_one = "10 - Hypoglycemic"
' 				If diet_msa_type_one = "11" Then diet_msa_type_one = "11 - Ketogenic"
' 				If diet_msa_type_one = "__" Then diet_msa_type_one = ""
'
' 				If diet_msa_type_two = "01" Then diet_msa_type_two = "01 - High Protein > 79 grams/day"
' 				If diet_msa_type_two = "02" Then diet_msa_type_two = "02 - Control Protein 40-60 grams/day"
' 				If diet_msa_type_two = "03" Then diet_msa_type_two = "03 - Control Protein < 40 grams/day"
' 				If diet_msa_type_two = "04" Then diet_msa_type_two = "04 - Lo Cholesterol"
' 				If diet_msa_type_two = "05" Then diet_msa_type_two = "05 - High Residue"
' 				If diet_msa_type_two = "06" Then diet_msa_type_two = "06 - Pregnancy and Lactation"
' 				If diet_msa_type_two = "07" Then diet_msa_type_two = "07 - Gluten Free"
' 				If diet_msa_type_two = "08" Then diet_msa_type_two = "08 - Lactose Free"
' 				If diet_msa_type_two = "09" Then diet_msa_type_two = "09 - Anti-Dumping"
' 				If diet_msa_type_two = "10" Then diet_msa_type_two = "10 - Hypoglycemic"
' 				If diet_msa_type_two = "11" Then diet_msa_type_two = "11 - Ketogenic"
' 				If diet_msa_type_two = "__" Then diet_msa_type_two = ""
'
' 				If diet_msa_type_three = "01" Then diet_msa_type_three = "01 - High Protein > 79 grams/day"
' 				If diet_msa_type_three = "02" Then diet_msa_type_three = "02 - Control Protein 40-60 grams/day"
' 				If diet_msa_type_three = "03" Then diet_msa_type_three = "03 - Control Protein < 40 grams/day"
' 				If diet_msa_type_three = "04" Then diet_msa_type_three = "04 - Lo Cholesterol"
' 				If diet_msa_type_three = "05" Then diet_msa_type_three = "05 - High Residue"
' 				If diet_msa_type_three = "06" Then diet_msa_type_three = "06 - Pregnancy and Lactation"
' 				If diet_msa_type_three = "07" Then diet_msa_type_three = "07 - Gluten Free"
' 				If diet_msa_type_three = "08" Then diet_msa_type_three = "08 - Lactose Free"
' 				If diet_msa_type_three = "09" Then diet_msa_type_three = "09 - Anti-Dumping"
' 				If diet_msa_type_three = "10" Then diet_msa_type_three = "10 - Hypoglycemic"
' 				If diet_msa_type_three = "11" Then diet_msa_type_three = "11 - Ketogenic"
' 				If diet_msa_type_three = "__" Then diet_msa_type_three = ""
'
' 				If diet_msa_type_four = "01" Then diet_msa_type_four = "01 - High Protein > 79 grams/day"
' 				If diet_msa_type_four = "02" Then diet_msa_type_four = "02 - Control Protein 40-60 grams/day"
' 				If diet_msa_type_four = "03" Then diet_msa_type_four = "03 - Control Protein < 40 grams/day"
' 				If diet_msa_type_four = "04" Then diet_msa_type_four = "04 - Lo Cholesterol"
' 				If diet_msa_type_four = "05" Then diet_msa_type_four = "05 - High Residue"
' 				If diet_msa_type_four = "06" Then diet_msa_type_four = "06 - Pregnancy and Lactation"
' 				If diet_msa_type_four = "07" Then diet_msa_type_four = "07 - Gluten Free"
' 				If diet_msa_type_four = "08" Then diet_msa_type_four = "08 - Lactose Free"
' 				If diet_msa_type_four = "09" Then diet_msa_type_four = "09 - Anti-Dumping"
' 				If diet_msa_type_four = "10" Then diet_msa_type_four = "10 - Hypoglycemic"
' 				If diet_msa_type_four = "11" Then diet_msa_type_four = "11 - Ketogenic"
' 				If diet_msa_type_four = "__" Then diet_msa_type_four = ""
'
' 				If diet_msa_type_five = "01" Then diet_msa_type_five = "01 - High Protein > 79 grams/day"
' 				If diet_msa_type_five = "02" Then diet_msa_type_five = "02 - Control Protein 40-60 grams/day"
' 				If diet_msa_type_five = "03" Then diet_msa_type_five = "03 - Control Protein < 40 grams/day"
' 				If diet_msa_type_five = "04" Then diet_msa_type_five = "04 - Lo Cholesterol"
' 				If diet_msa_type_five = "05" Then diet_msa_type_five = "05 - High Residue"
' 				If diet_msa_type_five = "06" Then diet_msa_type_five = "06 - Pregnancy and Lactation"
' 				If diet_msa_type_five = "07" Then diet_msa_type_five = "07 - Gluten Free"
' 				If diet_msa_type_five = "08" Then diet_msa_type_five = "08 - Lactose Free"
' 				If diet_msa_type_five = "09" Then diet_msa_type_five = "09 - Anti-Dumping"
' 				If diet_msa_type_five = "10" Then diet_msa_type_five = "10 - Hypoglycemic"
' 				If diet_msa_type_five = "11" Then diet_msa_type_five = "11 - Ketogenic"
' 				If diet_msa_type_five = "__" Then diet_msa_type_five = ""
'
' 				If diet_msa_type_six = "01" Then diet_msa_type_six = "01 - High Protein > 79 grams/day"
' 				If diet_msa_type_six = "02" Then diet_msa_type_six = "02 - Control Protein 40-60 grams/day"
' 				If diet_msa_type_six = "03" Then diet_msa_type_six = "03 - Control Protein < 40 grams/day"
' 				If diet_msa_type_six = "04" Then diet_msa_type_six = "04 - Lo Cholesterol"
' 				If diet_msa_type_six = "05" Then diet_msa_type_six = "05 - High Residue"
' 				If diet_msa_type_six = "06" Then diet_msa_type_six = "06 - Pregnancy and Lactation"
' 				If diet_msa_type_six = "07" Then diet_msa_type_six = "07 - Gluten Free"
' 				If diet_msa_type_six = "08" Then diet_msa_type_six = "08 - Lactose Free"
' 				If diet_msa_type_six = "09" Then diet_msa_type_six = "09 - Anti-Dumping"
' 				If diet_msa_type_six = "10" Then diet_msa_type_six = "10 - Hypoglycemic"
' 				If diet_msa_type_six = "11" Then diet_msa_type_six = "11 - Ketogenic"
' 				If diet_msa_type_six = "__" Then diet_msa_type_six = ""
'
' 				If diet_msa_type_seven = "01" Then diet_msa_type_seven = "01 - High Protein > 79 grams/day"
' 				If diet_msa_type_seven = "02" Then diet_msa_type_seven = "02 - Control Protein 40-60 grams/day"
' 				If diet_msa_type_seven = "03" Then diet_msa_type_seven = "03 - Control Protein < 40 grams/day"
' 				If diet_msa_type_seven = "04" Then diet_msa_type_seven = "04 - Lo Cholesterol"
' 				If diet_msa_type_seven = "05" Then diet_msa_type_seven = "05 - High Residue"
' 				If diet_msa_type_seven = "06" Then diet_msa_type_seven = "06 - Pregnancy and Lactation"
' 				If diet_msa_type_seven = "07" Then diet_msa_type_seven = "07 - Gluten Free"
' 				If diet_msa_type_seven = "08" Then diet_msa_type_seven = "08 - Lactose Free"
' 				If diet_msa_type_seven = "09" Then diet_msa_type_seven = "09 - Anti-Dumping"
' 				If diet_msa_type_seven = "10" Then diet_msa_type_seven = "10 - Hypoglycemic"
' 				If diet_msa_type_seven = "11" Then diet_msa_type_seven = "11 - Ketogenic"
' 				If diet_msa_type_seven = "__" Then diet_msa_type_seven = ""
'
' 				If diet_msa_type_eight = "01" Then diet_msa_type_eight = "01 - High Protein > 79 grams/day"
' 				If diet_msa_type_eight = "02" Then diet_msa_type_eight = "02 - Control Protein 40-60 grams/day"
' 				If diet_msa_type_eight = "03" Then diet_msa_type_eight = "03 - Control Protein < 40 grams/day"
' 				If diet_msa_type_eight = "04" Then diet_msa_type_eight = "04 - Lo Cholesterol"
' 				If diet_msa_type_eight = "05" Then diet_msa_type_eight = "05 - High Residue"
' 				If diet_msa_type_eight = "06" Then diet_msa_type_eight = "06 - Pregnancy and Lactation"
' 				If diet_msa_type_eight = "07" Then diet_msa_type_eight = "07 - Gluten Free"
' 				If diet_msa_type_eight = "08" Then diet_msa_type_eight = "08 - Lactose Free"
' 				If diet_msa_type_eight = "09" Then diet_msa_type_eight = "09 - Anti-Dumping"
' 				If diet_msa_type_eight = "10" Then diet_msa_type_eight = "10 - Hypoglycemic"
' 				If diet_msa_type_eight = "11" Then diet_msa_type_eight = "11 - Ketogenic"
' 				If diet_msa_type_eight = "__" Then diet_msa_type_eight = ""
'
' 				If diet_mf_verif_one = "_" Then diet_mf_verif_one = ""
' 				If diet_mf_verif_two = "_" Then diet_mf_verif_two = ""
' 				If diet_msa_verif_one = "_" Then diet_msa_verif_one = ""
' 				If diet_msa_verif_two = "_" Then diet_msa_verif_two = ""
' 				If diet_msa_verif_three = "_" Then diet_msa_verif_three = ""
' 				If diet_msa_verif_four = "_" Then diet_msa_verif_four = ""
' 				If diet_msa_verif_five = "_" Then diet_msa_verif_five = ""
' 				If diet_msa_verif_six = "_" Then diet_msa_verif_six = ""
' 				If diet_msa_verif_seven = "_" Then diet_msa_verif_seven = ""
' 				If diet_msa_verif_eight	 = "_" Then diet_msa_verif_eight = ""
' 			End If
' 		End If
' 	end sub
'
' 	public sub collect_parent_information()
'
' 		If pare_exists = TRUE Then
' 			' MsgBox "PARE EXISTS for " & ref_number
' 			pare_row_index = 0
' 			Do
' 				For the_membs = 0 to UBound(HH_MEMB_ARRAY)
' 					' MsgBox "REF on PARE - " & pare_child_ref_nbr(pare_row_index) & vbCr & "REF of the HH MEMB - " & HH_MEMB_ARRAY(the_membs).ref_number
' 					If pare_child_ref_nbr(pare_row_index) = HH_MEMB_ARRAY(the_membs).ref_number Then
' 						pare_child_member_index(pare_array_count) = the_membs
'
' 						If HH_MEMB_ARRAY(the_membs).parent_one_name = "" Then
'
' 							HH_MEMB_ARRAY(the_membs).parent_one_name = full_name_const
' 							HH_MEMB_ARRAY(the_membs).parent_one_type = pare_relationship_type(pare_array_count)
' 							HH_MEMB_ARRAY(the_membs).parent_one_verif = pare_verification(pare_array_count)
' 							HH_MEMB_ARRAY(the_membs).parent_one_in_home = TRUE
'
' 						ElseIf HH_MEMB_ARRAY(the_membs).parent_two_name = "" Then
' 							HH_MEMB_ARRAY(the_membs).parent_two_name = full_name_const
' 							HH_MEMB_ARRAY(the_membs).parent_two_type = pare_relationship_type(pare_array_count)
' 							HH_MEMB_ARRAY(the_membs).parent_two_verif = pare_verification(pare_array_count)
' 							HH_MEMB_ARRAY(the_membs).parent_two_in_home = TRUE
' 						End If
' 						' MsgBox HH_MEMB_ARRAY(the_membs).parent_one_name
'
' 						Exit For
' 					End If
' 				Next
' 				pare_row_index = pare_row_index + 1
' 			Loop until pare_row_index > UBound(pare_child_ref_nbr)
' 		End If
'
' 		Call navigate_to_MAXIS_screen("STAT", "ABPS")
' 		Do
' 			abps_row = 15
' 			Do
' 				EMReadScreen abps_ref_nrb, 2, abps_row, 35
' 				' MsgBox "REF on ABPS - " & abps_ref_nrb & vbCr & "REF of the HH MEMB - " & ref_number
' 				If abps_ref_nrb = ref_number Then
' 					EMReadScreen abps_last_name, 24, 10, 30
' 					EMReadScreen abps_first_name, 12, 10, 63
' 					EMReadScreen abps_mid_initial, 1, 10, 80
' 					EMReadScreen abps_ssn, 11, 11, 30
' 					EMReadScreen abps_dob, 10, 11, 60
' 					EMReadScreen abps_gender, 1, 11, 80
' 					EMReadScreen abps_parental_status, 1, abps_row, 53
' 					EMReadScreen abps_custody, 1, abps_row, 67
'
' 					abps_last_name = replace(abps_last_name, "_", "")
' 					abps_first_name = replace(abps_first_name, "_", "")
' 					abps_mid_initial = replace(abps_mid_initial, "_", "")
'
' 					' MsgBox trim(abps_first_name) & " " & trim(abps_last_name)
' 					If abps_first_name = "" AND abps_last_name = "" Then abps_first_name = "Name Unknown"
' 					abps_ssn = replace(abps_ssn, "_", "")
' 					abps_ssn = trim(abps_ssn)
' 					abps_ssn = replace(abps_ssn, " ", "-")
'
' 					abps_dob = replace(abps_dob, "_", "")
' 					abps_dob = trim(abps_dob)
' 					abps_dob = replace(abps_dob, " ", "/")
'
' 					If parent_one_name = "" Then
'
' 						parent_one_name = trim(abps_first_name) & " " & trim(abps_last_name)
' 						parent_one_type = "ABSENT"
' 						parent_one_verif = ""
' 						parent_one_in_home = FALSE
'
' 					ElseIf parent_two_name = "" Then
' 						parent_two_name = trim(abps_first_name) & " " & trim(abps_last_name)
' 						parent_two_type = "ABSENT"
' 						parent_two_verif = ""
' 						parent_two_in_home = FALSE
' 					End If
' 				End If
' 				abps_row = abps_row + 1
'
' 				If abps_row = 18 Then
' 					PF20
' 					abps_row = 15
' 					EMReadScreen end_of_list, 9, 24, 14
' 					If end_of_list = "LAST PAGE" Then Exit Do
' 				End If
' 			Loop until abps_ref_nrb = "__"
' 			transmit
' 			EMReadScreen last_abps, 7, 24, 2
' 		Loop until last_abps = "ENTER A"
'
'
' 	end sub
'
' 	Public sub choose_the_members()
'
' 	end sub
'
' 	' private sub Class_Initialize()
' 	' end sub
' end class
'
'
' class client_income
'
' 	'about the income
' 	public member_ref
' 	public member_name
' 	public member
' 	public access_denied
'
' 	public panel_name
' 	public panel_instance
'
' 	public unea_or_earned
' 	public income_type
' 	public income_type_code
' 	public income_review
' 	public income_verification
' 	public verif_explaination
' 	public income_start_date
' 	public income_end_date
' 	public pay_frequency
' 	public pay_weekday
' 	public hc_inc_est
' 	public most_recent_pay_date
' 	public most_recent_pay_amt
' 	public income_notes
' 	public pay_gross
' 	public expenses_allowed
' 	public expenses_not_allowed
'
' 	'JOBS
' 	public subsidized_income_type
' 	public hourly_wage
' 	public employer
' 	public prosp_pay_total
' 	public prosp_hours_total
' 	public prosp_pay_date_one
' 	public prosp_pay_wage_one
' 	public prosp_pay_date_two
' 	public prosp_pay_wage_two
' 	public prosp_pay_date_three
' 	public prosp_pay_wage_three
' 	public prosp_pay_date_four
' 	public prosp_pay_wage_four
' 	public prosp_pay_date_five
' 	public prosp_pay_wage_five
' 	public prosp_average_pay
'
' 	public retro_pay_total
' 	public retro_hours_total
' 	public retro_pay_date_one
' 	public retro_pay_wage_one
' 	public retro_pay_date_two
' 	public retro_pay_wage_two
' 	public retro_pay_date_three
' 	public retro_pay_wage_three
' 	public retro_pay_date_four
' 	public retro_pay_wage_four
' 	public retro_pay_date_five
' 	public retro_pay_wage_five
' 	public retro_average_pay
'
' 	'BUSI
' 	public prosp_net_cash_earnings
' 	public prosp_gross_cash_earnings
' 	public cash_earnings_verif
' 	public prosp_cash_expenses
' 	public cash_expense_verif
' 	public retro_net_cash_earnings
' 	public retro_gross_cash_earnings
' 	public retro_cash_expenses
'
' 	public prosp_net_ive_earnings
' 	public prosp_gross_ive_earnings
' 	public ive_earnings_verif
' 	public prosp_ive_expenses
' 	public ive_expense_verif
'
' 	public prosp_net_snap_earnings
' 	public prosp_gross_snap_earnings
' 	public snap_earnings_verif
' 	public prosp_snap_expenses
' 	public snap_expense_verif
' 	public retro_net_snap_earnings
' 	public retro_gross_snap_earnings
' 	public retro_snap_expenses
'
' 	public prosp_net_hc_a_earnings
' 	public prosp_gross_hc_a_earnings
' 	public hc_a_earnings_verif
' 	public prosp_hc_a_expenses
' 	public hc_a_expense_verif
'
' 	public prosp_net_hc_b_earnings
' 	public prosp_gross_hc_b_earnings
' 	public hc_b_earnings_verif
' 	public prosp_hc_b_expenses
' 	public hc_b_expense_verif
'
' 	public retro_reptd_hours
' 	public retro_min_wage_hours
' 	public prosp_reptd_hours
' 	public prosp_min_wage_hours
'
' 	public self_emp_method
' 	public self_emp_method_date
'
' 	'UNEA
' 	public claim_number
' 	public cola_month
'
'
' 	public sub read_member_name()
' 		Call navigate_to_MAXIS_screen("STAT", "MEMB")
' 		EMWriteScreen member_ref, 20, 76
' 		transmit
'
' 		EMReadScreen access_denied_check, 13, 24, 2         'Sometimes MEMB gets this access denied issue and we have to work around it.
' 		If access_denied_check = "ACCESS DENIED" Then
' 			PF10
' 			last_name = "UNABLE TO FIND"
' 			first_name = "Access Denied"
' 			access_denied = TRUE
' 		Else
' 			access_denied = FALSE
' 			EMReadscreen last_name, 25, 6, 30
' 			EMReadscreen first_name, 12, 6, 63
' 		End If
' 		last_name = trim(replace(last_name, "_", ""))
' 		first_name = trim(replace(first_name, "_", ""))
'
' 		member_name = first_name & " " & last_name
' 		member = member_ref & " - " & member_name
' 		' MsgBox "~" & member & "~"
' 	end sub
'
' 	Public sub read_jobs_panel()
' 		jobs_found = FALSE
' 	end sub
'
' 	Public sub read_busi_panel()
' 		busi_found = FALSE
' 	end sub
'
' 	Public sub read_unea_panel()
' 		Call navigate_to_MAXIS_screen("STAT", "UNEA")
' 		EMWriteScreen member_ref, 20, 76
' 		EMWriteScreen panel_instance, 20, 79
' 		transmit
'
' 		panel_name = "UNEA"
' 		unea_or_earned = "Unearned"
'
' 		EMReadScreen income_type, 2, 5, 37
' 		EMReadScreen income_verification, 1, 5, 65
' 		EMReadScreen income_start_date, 8, 7, 37
' 		EMReadScreen income_end_date, 8, 7, 68
'
' 		EmWriteScreen "X", 6, 56
' 		transmit
' 			EMReadScreen pay_frequency, 1, 10, 63
' 			EMReadScreen hc_inc_est, 8, 9, 65
' 		PF3
'
' 		EMReadScreen claim_number, 15, 6, 37
' 		EMReadScreen cola_month, 2, 19, 36
'
' 		EMReadScreen prosp_pay_total, 8, 18, 68
' 		EMReadScreen prosp_pay_date_one, 8, 13, 54
' 		EMReadScreen prosp_pay_wage_one, 8, 13, 68
' 		EMReadScreen prosp_pay_date_two, 8, 14, 54
' 		EMReadScreen prosp_pay_wage_two, 8, 14, 68
' 		EMReadScreen prosp_pay_date_three, 8, 15, 54
' 		EMReadScreen prosp_pay_wage_three, 8, 15, 68
' 		EMReadScreen prosp_pay_date_four, 8, 16, 54
' 		EMReadScreen prosp_pay_wage_four, 8, 16, 68
' 		EMReadScreen prosp_pay_date_five, 8, 17, 54
' 		EMReadScreen prosp_pay_wage_five, 8, 17, 68
'
' 		EMReadScreen retro_pay_total, 8, 18, 39
' 		EMReadScreen retro_pay_date_one, 8, 13, 25
' 		EMReadScreen retro_pay_wage_one, 8, 13, 39
' 		EMReadScreen retro_pay_date_two, 8, 14, 25
' 		EMReadScreen retro_pay_wage_two, 8, 14, 39
' 		EMReadScreen retro_pay_date_three, 8, 15, 25
' 		EMReadScreen retro_pay_wage_three, 8, 15, 39
' 		EMReadScreen retro_pay_date_four, 8, 16, 25
' 		EMReadScreen retro_pay_wage_four, 8, 16, 39
' 		EMReadScreen retro_pay_date_five, 8, 17, 25
' 		EMReadScreen retro_pay_wage_five, 8, 17, 39
'
' 		income_type_code = income_type
' 		If income_type = "01" Then income_type = "01 - RSDI, Disa"
' 		If income_type = "02" Then income_type = "02 - RSDI, No Disa"
' 		If income_type = "03" Then income_type = "03 - SSI"
' 		If income_type = "06" Then income_type = "06 - Non-MN PA"
' 		If income_type = "11" Then income_type = "11 - VA Disability Benefit"
' 		If income_type = "12" Then income_type = "12 - VA Pension"
' 		If income_type = "13" Then income_type = "13 - VA Other"
' 		If income_type = "38" Then income_type = "38 - VA Aid and Attendance"
' 		If income_type = "14" Then income_type = "14 - Unemployment Insurance"
' 		If income_type = "15" Then income_type = "15 - Worker's Compensation"
' 		If income_type = "16" Then income_type = "16 - Railroad Retirement"
' 		If income_type = "17" Then income_type = "17 - Other Retirement"
' 		If income_type = "18" Then income_type = "18 - Military Entitlement"
' 		If income_type = "19" Then income_type = "19 - FC Child Requesting SNAP"
' 		If income_type = "20" Then income_type = "20 - FC Child NOT Requesting SNAP"
' 		If income_type = "21" Then income_type = "21 - FC Adult Requesting SNAP"
' 		If income_type = "22" Then income_type = "22 - FC Adult NOT Requesting SNAP"
' 		If income_type = "23" Then income_type = "23 - Dividends"
' 		If income_type = "24" Then income_type = "24 - Interest "
' 		If income_type = "25" Then income_type = "25 - Counted Gifts or Prizes"
' 		If income_type = "26" Then income_type = "26 - Strike Benefit"
' 		If income_type = "27" Then income_type = "27 - Contract for Deed"
' 		If income_type = "28" Then income_type = "28 - Illegal Income"
' 		If income_type = "29" Then income_type = "29 - Other Countable"
' 		If income_type = "30" Then income_type = "30 - Not Counted - Infreq <30"
' 		If income_type = "21" Then income_type = "31 - Other SNAP Only"
' 		If income_type = "08" Then income_type = "08 - Direct Child Support"
' 		If income_type = "35" Then income_type = "35 - Direct Spousal Support"
' 		If income_type = "36" Then income_type = "36 - Disb Child Support"
' 		If income_type = "37" Then income_type = "37 - Disb Spousal Support"
' 		If income_type = "39" Then income_type = "39 - Disb Child Support Arrears"
' 		If income_type = "40" Then income_type = "40 - Disb Spousal Support Arrears"
' 		If income_type = "43" Then income_type = "43 - Disb Excess Child Support"
' 		If income_type = "44" Then income_type = "44 - MSA - Excess Income for SSI"
' 		If income_type = "45" Then income_type = "45 - County 88 Child Support"
' 		If income_type = "46" Then income_type = "46 - County 88 Gaming"
' 		If income_type = "47" Then income_type = "47 - Counted Tribal Income"
' 		If income_type = "48" Then income_type = "48 - Trust Income"
' 		If income_type = "49" Then income_type = "49 - Non-Recurring > $60/qtr"
'
' 		If income_verification = "1" Then income_verification = "1 - Copy of Checks"
' 		If income_verification = "2" Then income_verification = "2 - Award Letters"
' 		If income_verification = "3" Then income_verification = "3 - System Initiated"
' 		If income_verification = "4" Then income_verification = "4 - Colateral Statement"
' 		If income_verification = "5" Then income_verification = "5 - Pend Out State Verif"
' 		If income_verification = "6" Then income_verification = "6 - Other Document"
' 		If income_verification = "7" Then income_verification = "7 - Worker Initiated"
' 		If income_verification = "8" Then income_verification = "8 - RI Stubs"
' 		If income_verification = "N" Then income_verification = "N - No Verif Provided"
' 		' MsgBox "~" & income_verification & "~"
' 		income_start_date = replace(income_start_date, " ", "/")
' 		If income_start_date = "__/__/__" Then income_start_date = ""
' 		income_end_date = replace(income_end_date, " ", "/")
' 		If income_end_date = "__/__/__" Then income_end_date = ""
'
' 		If pay_frequency = "1" Then pay_frequency = "1 - Monthly"
' 		If pay_frequency = "2" Then pay_frequency = "2 - Semi-monthly"
' 		If pay_frequency = "3" Then pay_frequency = "3 - Biweekly"
' 		If pay_frequency = "4" Then pay_frequency = "4 - Weekly"
' 		If pay_frequency = "5" Then pay_frequency = "5 - Other"
' 		If pay_frequency = "_" Then pay_frequency = ""
' 		hc_inc_est = trim(hc_inc_est)
'
' 		'pay_weekday'
'
' 		claim_number = replace(claim_number, "_", "")
'
' 		If cola_month = "01" Then cola_month = "January"
' 		If cola_month = "02" Then cola_month = "February"
' 		If cola_month = "03" Then cola_month = "March"
' 		If cola_month = "04" Then cola_month = "April"
' 		If cola_month = "05" Then cola_month = "May"
' 		If cola_month = "06" Then cola_month = "June"
' 		If cola_month = "07" Then cola_month = "July"
' 		If cola_month = "08" Then cola_month = "August"
' 		If cola_month = "09" Then cola_month = "September"
' 		If cola_month = "10" Then cola_month = "October"
' 		If cola_month = "11" Then cola_month = "November"
' 		If cola_month = "12" Then cola_month = "December"
' 		If cola_month = "NA" Then cola_month = "Not Applicable"
' 		If cola_month = "__" Then cola_month = "Unspecified"
'
' 		prosp_pay_total = trim(prosp_pay_total)
' 		prosp_pay_date_one = replace(prosp_pay_date_one, " ", "/")
' 		If prosp_pay_date_one = "__/__/__" Then prosp_pay_date_one = ""
' 		prosp_pay_wage_one = trim(prosp_pay_wage_one)
' 		If prosp_pay_wage_one = "________" Then prosp_pay_wage_one = ""
' 		prosp_pay_date_two = replace(prosp_pay_date_two, " ", "/")
' 		If prosp_pay_date_two = "__/__/__" Then prosp_pay_date_two = ""
' 		prosp_pay_wage_two = trim(prosp_pay_wage_two)
' 		If prosp_pay_wage_two = "________" Then prosp_pay_wage_two = ""
' 		prosp_pay_date_three = replace(prosp_pay_date_three, " ", "/")
' 		If prosp_pay_date_three = "__/__/__" Then prosp_pay_date_three = ""
' 		prosp_pay_wage_three = trim(prosp_pay_wage_three)
' 		If prosp_pay_wage_three = "________" Then prosp_pay_wage_three = ""
' 		prosp_pay_date_four = replace(prosp_pay_date_four, " ", "/")
' 		If prosp_pay_date_four = "__/__/__" Then prosp_pay_date_four = ""
' 		prosp_pay_wage_four = trim(prosp_pay_wage_four)
' 		If prosp_pay_wage_four = "________" Then prosp_pay_wage_four = ""
' 		prosp_pay_date_five = replace(prosp_pay_date_five, " ", "/")
' 		If prosp_pay_date_five = "__/__/__" Then prosp_pay_date_five = ""
' 		prosp_pay_wage_five = trim(prosp_pay_wage_five)
' 		If prosp_pay_wage_five = "________" Then prosp_pay_wage_five = ""
' 		total_of_prosp_pay = 0
' 		number_of_checks = 0
' 		If prosp_pay_wage_one <> "" Then
' 			total_of_prosp_pay = total_of_prosp_pay + prosp_pay_wage_one * 1
' 			number_of_checks = number_of_checks + 1
' 		End If
' 		If prosp_pay_wage_two <> "" Then
' 			total_of_prosp_pay = total_of_prosp_pay + prosp_pay_wage_two * 1
' 			number_of_checks = number_of_checks + 1
' 		End If
' 		If prosp_pay_wage_three <> "" Then
' 			total_of_prosp_pay = total_of_prosp_pay + prosp_pay_wage_three * 1
' 			number_of_checks = number_of_checks + 1
' 		End If
' 		If prosp_pay_wage_four <> "" Then
' 			total_of_prosp_pay = total_of_prosp_pay + prosp_pay_wage_four * 1
' 			number_of_checks = number_of_checks + 1
' 		End If
' 		If prosp_pay_wage_five <> "" Then
' 			total_of_prosp_pay = total_of_prosp_pay + prosp_pay_wage_five * 1
' 			number_of_checks = number_of_checks + 1
' 		End If
' 		If number_of_checks <> 0 Then prosp_average_pay = total_of_prosp_pay / number_of_checks
' 		prosp_average_pay = prosp_average_pay & ""
'
' 		retro_pay_total = trim(retro_pay_total)
' 		retro_pay_date_one = replace(retro_pay_date_one, " ", "/")
' 		If retro_pay_date_one = "__/__/__" Then retro_pay_date_one = ""
' 		retro_pay_wage_one = trim(retro_pay_wage_one)
' 		If retro_pay_wage_one = "________" Then retro_pay_wage_one = ""
' 		retro_pay_date_two = replace(retro_pay_date_two, " ", "/")
' 		If retro_pay_date_two = "__/__/__" Then retro_pay_date_two = ""
' 		retro_pay_wage_two = trim(retro_pay_wage_two)
' 		If retro_pay_wage_two = "________" Then retro_pay_wage_two = ""
' 		retro_pay_date_three = replace(retro_pay_date_three, " ", "/")
' 		If retro_pay_date_three = "__/__/__" Then retro_pay_date_three = ""
' 		retro_pay_wage_three = trim(retro_pay_wage_three)
' 		If retro_pay_wage_three = "________" Then retro_pay_wage_three = ""
' 		retro_pay_date_four = replace(retro_pay_date_four, " ", "/")
' 		If retro_pay_date_four = "__/__/__" Then retro_pay_date_four = ""
' 		retro_pay_wage_four = trim(retro_pay_wage_four)
' 		If retro_pay_wage_four = "________" Then retro_pay_wage_four = ""
' 		retro_pay_date_five = replace(retro_pay_date_five, " ", "/")
' 		If retro_pay_date_five = "__/__/__" Then retro_pay_date_five = ""
' 		retro_pay_wage_five = trim(retro_pay_wage_five)
' 		If retro_pay_wage_five = "________" Then retro_pay_wage_five = ""
' 		total_of_retro_pay = 0
' 		number_of_checks = 0
' 		If retro_pay_wage_one <> "" Then
' 			total_of_retro_pay = total_of_retro_pay + retro_pay_wage_one * 1
' 			number_of_checks = number_of_checks + 1
' 		End If
' 		If retro_pay_wage_two <> "" Then
' 			total_of_retro_pay = total_of_retro_pay + retro_pay_wage_two * 1
' 			number_of_checks = number_of_checks + 1
' 		End If
' 		If retro_pay_wage_three <> "" Then
' 			total_of_retro_pay = total_of_retro_pay + retro_pay_wage_three * 1
' 			number_of_checks = number_of_checks + 1
' 		End If
' 		If retro_pay_wage_four <> "" Then
' 			total_of_retro_pay = total_of_retro_pay + retro_pay_wage_four * 1
' 			number_of_checks = number_of_checks + 1
' 		End If
' 		If retro_pay_wage_five <> "" Then
' 			total_of_retro_pay = total_of_retro_pay + retro_pay_wage_five * 1
' 			number_of_checks = number_of_checks + 1
' 		End If
' 		If number_of_checks <> 0 Then retro_average_pay = total_of_retro_pay / number_of_checks
' 		retro_average_pay = retro_average_pay & ""
'
' 		If pay_frequency = "3 - Biweekly" OR pay_frequency = "4 - Weekly" Then
' 			If prosp_pay_date_five <> "" Then
' 				pay_weekday = WeekdayName(weekday(prosp_pay_date_five))
' 			ElseIf prosp_pay_date_four <> "" Then
' 				pay_weekday = WeekdayName(weekday(prosp_pay_date_four))
' 			ElseIf prosp_pay_date_three <> "" Then
' 				pay_weekday = WeekdayName(weekday(prosp_pay_date_three))
' 			ElseIf prosp_pay_date_two <> "" Then
' 				pay_weekday = WeekdayName(weekday(prosp_pay_date_two))
' 			ElseIf prosp_pay_date_one <> "" Then
' 				pay_weekday = WeekdayName(weekday(prosp_pay_date_one))
' 			End If
'
' 		End If
'
' 	end sub
'
' end class
'
' class client_assets
'
' 	public member_ref
' 	public member_name
' 	public member
' 	public access_denied
'
' 	public panel_name
' 	public panel_instance
'
' 	public asset_btn_one
' 	public asset_type
' 	public account_number
' 	public asset_verification
' 	public asset_update_date
' 	public withdraw_yn
' 	public withdraw_penalty
' 	public withdraw_verif
' 	public count_cash_yn
' 	public count_snap_yn
' 	public count_hc_yn
' 	public count_grh_yn
' 	public count_ive_yn
' 	public joint_owners_yn
' 	public share_ratio
' 	public next_interest_date
'
' 	public cash_value
'
' 	public acct_location
' 	public acct_balance
' 	public acct_balance_date
'
' 	public cars_year
' 	public cars_make
' 	public cars_model
' 	public cars_trade_in_value
' 	public cars_loan_value
' 	public cars_value_source
' 	public cars_amt_owed
' 	public cars_owed_verification
' 	public cars_owed_date
' 	public cars_use
' 	public cars_hc_benefit
'
' 	public secu_name
' 	public secu_cash_value
' 	public secu_cash_value_date
' 	public secu_face_value
'
' 	public rest_market_value
' 	public rest_value_verification
' 	public rest_amount_owed
' 	public rest_owed_verification
' 	public rest_owed_date
' 	public rest_property_status
' 	public rest_ive_repayment_agreement_date
'
' 	' function access_ACCT_panel(access_type, member_name,
'
' 	' account_type,
' 	' account_number,
' 	' account_location,
' 	' account_balance,
' 	' account_verification,
' 	' update_date, panel_ref_numb,
' 	' balance_date,
' 	' withdraw_penalty,
' 	' withdraw_yn,
' 	' withdraw_verif_code,
' 	' count_cash,
' 	' count_snap,
' 	' count_hc,
' 	' count_grh,
' 	' count_ive,
' 	' joint_own_yn,
' 	' share_ratio,
' 	' next_interest)
'
' 	' function access_CARS_panel(access_type, member_name,
'
' 	' cars_type,
' 	' cars_year,
' 	' cars_make,
' 	' cars_model,
' 	' cars_verif,
' 	' update_date, panel_ref_numb,
' 	' cars_trade_in,
' 	' cars_loan,
' 	' cars_source,
' 	' cars_owed,
' 	' cars_owed_verif_code,
' 	' cars_owed_date,
' 	' cars_use,
' 	' cars_hc_benefit,
' 	' cars_joint_yn,
' 	' cars_share)
'
' 	' function access_SECU_panel(access_type, member_name,
'
' 	' security_type,
' 	' security_account_number,
' 	' security_name,
' 	' security_cash_value,
' 	' security_verif,
' 	' secu_update_date,
' 	' panel_ref_numb,
' 	' security_face_value,
' 	' security_withdraw,
' 	' security_withdraw_yn,
' 	' security_withdraw_verif,
' 	' secu_cash_yn,
' 	' secu_snap_yn,
' 	' secu_hc_yn,
' 	' secu_grh_yn,
' 	' secu_ive_yn,
' 	' secu_joint,
' 	' secu_ratio,
' 	' security_eff_date)
'
' 	' function access_REST_panel(access_type, member_name,
'
' 	' rest_type,
' 	' rest_verif,
' 	' rest_update_date,
' 	' panel_ref_numb,
' 	' rest_market_value,
' 	' value_verif_code,
' 	' rest_amt_owed,
' 	' amt_owed_verif_code,
' 	' rest_eff_date,
' 	' rest_status,
' 	' rest_joint_yn,
' 	' rest_ratio,
' 	' repymt_agree_date)
'
' 	public sub read_member_name()
' 		Call navigate_to_MAXIS_screen("STAT", "MEMB")
' 		EMWriteScreen member_ref, 20, 76
' 		transmit
'
' 		EMReadScreen access_denied_check, 13, 24, 2         'Sometimes MEMB gets this access denied issue and we have to work around it.
' 		If access_denied_check = "ACCESS DENIED" Then
' 			PF10
' 			last_name = "UNABLE TO FIND"
' 			first_name = "Access Denied"
' 			access_denied = TRUE
' 		Else
' 			access_denied = FALSE
' 			EMReadscreen last_name, 25, 6, 30
' 			EMReadscreen first_name, 12, 6, 63
' 		End If
' 		last_name = trim(replace(last_name, "_", ""))
' 		first_name = trim(replace(first_name, "_", ""))
'
' 		member_name = first_name & " " & last_name
' 		member = member_ref & " - " & member_name
' 		' MsgBox "~" & member & "~"
' 	end sub
'
' 	public sub read_cash_panel()
' 		Call navigate_to_MAXIS_screen("STAT", "CASH")
' 		EMWriteScreen member_ref, 20, 76
' 		EMWriteScreen panel_instance, 20, 79
' 		transmit
'
' 		asset_type = "CASH"
'
' 		EMReadScreen cash_value, 8, 8, 39
' 		cash_value = trim(cash_value)
'
' 	end sub
'
' 	public sub read_acct_panel()
'
' 		Call navigate_to_MAXIS_screen("STAT", "ACCT")
' 		EMWriteScreen member_ref, 20, 76
' 		EMWriteScreen panel_instance, 20, 79
' 		transmit
'
' 		EMReadScreen panel_type, 2, 6, 44
' 		EMReadScreen panel_number, 20, 7, 44
' 		EMReadScreen acct_location, 20, 8, 44
' 		EMReadScreen panel_balance, 8, 10, 46
' 		EMReadScreen panel_verif_code, 1, 10, 64
' 		EMReadScreen balance_date, 8, 11, 44
' 		EMReadScreen withdraw_penalty, 8, 12, 46
' 		EMReadScreen withdraw_yn, 1, 12, 64
' 		EMReadScreen withdraw_verif_code, 1, 12, 72
' 		EMReadScreen count_cash, 1, 14, 50
' 		EMReadScreen count_snap, 1, 14, 57
' 		EMReadScreen count_hc, 1, 14, 64
' 		EMReadScreen count_grh, 1, 14, 72
' 		EMReadScreen count_ive, 1, 14, 80
' 		EMReadScreen joint_own_yn, 1, 15, 44
' 		EMReadScreen share_ratio, 5, 15, 76
' 		EMReadScreen next_interest, 5, 17, 57
' 		EMReadScreen update_date, 8, 21, 55
'
' 		If panel_type = "SV" Then asset_type = "SV - Savings"
' 		If panel_type = "CK" Then asset_type = "CK - Checking"
' 		If panel_type = "CE" Then asset_type = "CE - Certificate of Deposit"
' 		If panel_type = "MM" Then asset_type = "MM - Money Market"
' 		If panel_type = "DC" Then asset_type = "DC - Debit Card"
' 		If panel_type = "KO" Then asset_type = "KO - Keogh Account"
' 		If panel_type = "FT" Then asset_type = "FT - Fed Thrift Savings Plan"
' 		If panel_type = "SL" Then asset_type = "SL - State & Local Govt"
' 		If panel_type = "RA" Then asset_type = "RA - Employee Ret Annuities"
' 		If panel_type = "NP" Then asset_type = "NP - Non-Profit Emmployee Ret"
' 		If panel_type = "IR" Then asset_type = "IR - Indiv Ret Acct"
' 		If panel_type = "RH" Then asset_type = "RH - Roth IRA"
' 		If panel_type = "FR" Then asset_type = "FR - Ret Plan for Employers"
' 		If panel_type = "CT" Then asset_type = "CT - Corp Ret Trust"
' 		If panel_type = "RT" Then asset_type = "RT - Other Ret Fund"
' 		If panel_type = "QT" Then asset_type = "QT - Qualified Tuition (529)"
' 		If panel_type = "CA" Then asset_type = "CA - Coverdell SV (530)"
' 		If panel_type = "OE" Then asset_type = "OE - Other Educational"
' 		If panel_type = "OT" Then asset_type = "OT - Other"
'
' 		account_number = replace(panel_number, "_", "")
' 		acct_location =  replace(acct_location, "_", "")
' 		acct_balance = trim(panel_balance)
'
' 		If panel_verif_code = "1"  Then asset_verification = "1 - Bank Statement"
' 		If panel_verif_code = "2"  Then asset_verification = "2 - Agcy Ver Form"
' 		If panel_verif_code = "3"  Then asset_verification = "3 - Coltrl Contact"
' 		If panel_verif_code = "5"  Then asset_verification = "5 - Other Document"
' 		If panel_verif_code = "6"  Then asset_verification = "6 - Personal Statement"
' 		If panel_verif_code = "N"  Then asset_verification = "N - No Ver Prvd"
'
' 		acct_balance_date = replace(balance_date, " ", "/")
' 		If acct_balance_date = "__/__/__" Then acct_balance_date = ""
'
' 		withdraw_penalty = replace(withdraw_penalty, "_", "")
' 		withdraw_penalty = trim(withdraw_penalty)
' 		withdraw_yn = replace(withdraw_yn, "_", "")
' 		If withdraw_verif_code = "1"  Then withdraw_verif = "1 - Bank Statement"
' 		If withdraw_verif_code = "2"  Then withdraw_verif = "2 - Agcy Ver Form"
' 		If withdraw_verif_code = "3"  Then withdraw_verif = "3 - Coltrl Contact"
' 		If withdraw_verif_code = "5"  Then withdraw_verif = "5 - Other Document"
' 		If withdraw_verif_code = "6"  Then withdraw_verif = "6 - Personal Statement"
' 		If withdraw_verif_code = "N"  Then withdraw_verif = "N - No Ver Prvd"
'
' 		count_cash_yn = replace(count_cash, "_", "")
' 		count_snap_yn = replace(count_snap, "_", "")
' 		count_hc_yn = replace(count_hc, "_", "")
' 		count_grh_yn = replace(count_grh, "_", "")
' 		count_ive_yn = replace(count_ive, "_", "")
'
' 		share_ratio = replace(share_ratio, " ", "")
'
' 		next_interest_date = replace(next_interest, " ", "/")
' 		If next_interest_date = "__/__" Then next_interest_date = ""
'
' 		asset_update_date = replace(update_date, " ", "/")
'
' 	end sub
'
' 	public sub read_secu_panel()
' 		Call navigate_to_MAXIS_screen("STAT", "SECU")
' 		EMWriteScreen member_ref, 20, 76
' 		EMWriteScreen panel_instance, 20, 79
' 		transmit
'
'         EMReadScreen panel_type, 2, 6, 50
'         EMReadScreen security_account_number, 12, 7, 50
'         EMReadScreen security_name, 20, 8, 50
'         EMReadScreen security_cash_value, 8, 10, 52
'         EMReadScreen security_eff_date, 8, 11, 35   'not output
'         EMReadScreen verif_code, 1, 11, 50
'         EMReadScreen security_face_value, 8, 12, 52     'not output
'         EMReadScreen security_withdraw, 8, 13, 52       'not output
'         EMReadScreen security_withdraw_yn, 1, 13, 72    'not output
'         EMReadScreen security_withdraw_verif, 1, 13, 80 'not output
'
'         EMReadScreen secu_cash_yn, 1, 15, 50    'not output
'         EMReadScreen secu_snap_yn, 1, 15, 57    'not output
'         EMReadScreen secu_hc_yn, 1, 15, 64      'not output
'         EMReadScreen secu_grh_yn, 1, 15, 72     'not output
'         EMReadScreen secu_ive_yn, 1, 15, 80     'not output
'
'         EMReadScreen secu_joint, 1, 16, 44      'not output
'         EMReadScreen secu_ratio, 5, 16, 76      'not output
'         EMReadScreen secu_update_date, 8, 21, 55
'
'         If panel_type = "LI" Then asset_type = "LI - Life Insurance"
'         If panel_type = "ST" Then asset_type = "ST - Stocks"
'         If panel_type = "BO" Then asset_type = "BO - Bonds"
'         If panel_type = "CD" Then asset_type = "CD - Ctrct for Deed"
'         If panel_type = "MO" Then asset_type = "MO - Mortgage Note"
'         If panel_type = "AN" Then asset_type = "AN - Annuity"
'         If panel_type = "OT" Then asset_type = "OT - Other"
'
'         account_number = replace(security_account_number, "_", "")
'         secu_name = replace(security_name, "_", "")
'
'         secu_cash_value = replace(security_cash_value, "_", "")
'         secu_cash_value = trim(secu_cash_value)
'
'         secu_cash_value_date = replace(security_eff_date, " ", "/")
'         If secu_cash_value_date = "__/__/__" Then secu_cash_value_date = ""
'
'         If verif_code = "1" Then asset_verification = "1 - Agency Form"
'         If verif_code = "2" Then asset_verification = "2 - Source Doc"
'         If verif_code = "3" Then asset_verification = "3 - Phone Contact"
'         If verif_code = "5" Then asset_verification = "5 - Other Document"
'         If verif_code = "6" Then asset_verification = "6 - Personal Statement"
'         If verif_code = "N" Then asset_verification = "N - No Ver Prov"
'
'         secu_face_value = replace(security_face_value, "_", "")
'         secu_face_value = trim(secu_face_value)
'
'         withdraw_penalty = replace(security_withdraw, "_", "")
'         withdraw_penalty = trim(withdraw_penalty)
'
'         withdraw_yn = replace(security_withdraw_yn, "_", "")
'
'         If security_withdraw_verif = "1" Then withdraw_verif = "1 - Agency Form"
'         If security_withdraw_verif = "2" Then withdraw_verif = "2 - Source Doc"
'         If security_withdraw_verif = "3" Then withdraw_verif = "3 - Phone Contact"
'         If security_withdraw_verif = "4" Then withdraw_verif = "4 - Other Document"
'         If security_withdraw_verif = "5" Then withdraw_verif = "5 - Personal Stmt"
'         If security_withdraw_verif = "N" Then withdraw_verif = "N - No Ver Prov"
'
'         count_cash_yn = replace(secu_cash_yn, "_", "")
'         count_snap_yn = replace(secu_snap_yn, "_", "")
'         count_hc_yn = replace(secu_hc_yn, "_", "")
'         count_grh_yn = replace(secu_grh_yn, "_", "")
'         count_ive_yn = replace(secu_ive_yn, "_", "")
'
'         joint_owners_yn = replace(secu_joint, "_", "")
'         share_ratio = replace(secu_ratio, " ", "")
'
'         asset_update_date = replace(secu_update_date, " ", "/")
'
' 	end sub
'
' 	public sub read_cars_panel()
' 		Call navigate_to_MAXIS_screen("STAT", "CARS")
' 		EMWriteScreen member_ref, 20, 76
' 		EMWriteScreen panel_instance, 20, 79
' 		transmit
'
' 		EMReadScreen cars_type, 1, 6, 43
' 		EMReadScreen cars_year, 4, 8, 31
' 		EMReadScreen cars_make, 15, 8, 43
' 		EMReadScreen cars_model, 15, 8, 66
' 		EMReadScreen cars_trade_in, 8, 9, 45            'not output
' 		EMReadScreen cars_loan, 8, 9, 62                'not output
' 		EMReadScreen cars_source, 1, 9, 80              'not output
' 		EMReadScreen cars_verif_code, 1, 10, 60
' 		EMReadScreen cars_owed, 8, 12, 45               'not output
' 		EMReadScreen cars_owed_verif_code, 1, 12, 60    'not output
' 		EMReadScreen cars_owed_date, 8, 13, 43          'not output
' 		EMReadScreen cars_use, 1, 15, 43                'not output
' 		EMReadScreen cars_hc_benefit, 1, 15, 76         'not output
' 		EMReadScreen cars_joint_yn, 1, 16, 43           'not output
' 		EMReadScreen cars_share, 5, 16, 76              'not output
' 		EMReadScreen cars_update, 8, 21, 55
'
' 		If cars_type = "1" Then asset_type = "1 - Car"
' 		If cars_type = "2" Then asset_type = "2 - Truck"
' 		If cars_type = "3" Then asset_type = "3 - Van"
' 		If cars_type = "4" Then asset_type = "4 - Camper"
' 		If cars_type = "5" Then asset_type = "5 - Motorcycle"
' 		If cars_type = "6" Then asset_type = "6 - Trailer"
' 		If cars_type = "7" Then asset_type = "7 - Other"
'
' 		cars_make = replace(cars_make, "_", "")
' 		cars_model = replace(cars_model, "_", "")
'
' 		cars_trade_in_value = replace(cars_trade_in, "_", "")
' 		cars_trade_in_value = trim(cars_trade_in_value)
'
' 		cars_loan_value = replace(cars_loan, "_", "")
' 		cars_loan_value = trim(cars_loan_value)
'
' 		If cars_source = "1" Then cars_value_source = "1 - NADA"
' 		If cars_source = "2" Then cars_value_source = "2 - Appraisal Val"
' 		If cars_source = "3" Then cars_value_source = "3 - Client Stmt"
' 		If cars_source = "4" Then cars_value_source = "4 - Other Document"
'
' 		If cars_verif_code = "1" Then asset_verification = "1 - Title"
' 		If cars_verif_code = "2" Then asset_verification = "2 - License Reg"
' 		If cars_verif_code = "3" Then asset_verification = "3 - DMV"
' 		If cars_verif_code = "4" Then asset_verification = "4 - Purchase Agmt"
' 		If cars_verif_code = "5" Then asset_verification = "5 - Other Document"
' 		If cars_verif_code = "N" Then asset_verification = "N - No Ver Prvd"
'
' 		cars_amt_owed = replace(cars_owed, "_", "")
' 		cars_amt_owed = trim(cars_amt_owed)
'
' 		If cars_owed_verif_code = "1" Then cars_owed_verification = "1 - Bank/Lending Inst Stmt"
' 		If cars_owed_verif_code = "2" Then cars_owed_verification = "2 - Private Lender Stmt"
' 		If cars_owed_verif_code = "3" Then cars_owed_verification = "3 - Other Document"
' 		If cars_owed_verif_code = "4" Then cars_owed_verification = "4 - Pend Out State Verif"
' 		If cars_owed_verif_code = "N" Then cars_owed_verification = "N - No Ver Prvd"
'
' 		cars_owed_date = replace(cars_owed_date, " ", "/")
' 		If cars_owed_date = "__/__/__" Then cars_owed_date = ""
'
' 		If cars_use = "1" Then cars_use = "1 - Primary Vehicle"
' 		If cars_use = "2" Then cars_use = "2 - Employment/Training Search"
' 		If cars_use = "3" Then cars_use = "3 - Disa Transportation"
' 		If cars_use = "4" Then cars_use = "4 - Income Producing"
' 		If cars_use = "5" Then cars_use = "5 - Used as Home"
' 		If cars_use = "7" Then cars_use = "7 - Unlicensed"
' 		If cars_use = "8" Then cars_use = "8 - Other Countable"
' 		If cars_use = "9" Then cars_use = "9 - Unavailable"
' 		If cars_use = "0" Then cars_use = "0 - Long Distance Employment Travel"
' 		If cars_use = "A" Then cars_use = "A - Carry Heating Fuel or Water"
'
' 		cars_hc_benefit = replace(cars_hc_benefit, "_", "")
' 		joint_owners_yn = replace(cars_joint_yn, "_", "")
' 		share_ratio = replace(cars_share, " ", "")
'
' 		asset_update_date = replace(cars_update, " ", "/")
'
' 	end sub
'
' 	public sub read_rest_panel()
' 		Call navigate_to_MAXIS_screen("STAT", "REST")
' 		EMWriteScreen member_ref, 20, 76
' 		EMWriteScreen panel_instance, 20, 79
' 		transmit
'
'         EMReadScreen type_code, 1, 6, 39
'         EMReadScreen type_verif_code, 2, 6, 62
'         EMReadScreen rest_market_value, 10, 8, 41
'         EMReadScreen value_verif_code, 2, 8, 62
'         EMReadScreen rest_amt_owed, 10, 9, 41
'         EMReadScreen amt_owed_verif_code, 2, 9, 62
'         EMReadScreen rest_eff_date, 8, 10, 39
'         EMReadScreen rest_status, 1, 12, 54
'         EMReadScreen rest_joint_yn, 1, 13, 54
'         EMReadScreen rest_ratio, 5, 14, 54
'         EMReadScreen repymt_agree_date, 8, 16, 62
'         EMReadScreen rest_update_date, 8, 21, 55
'
'         If type_code = "1" Then asset_type = "1 - House"
'         If type_code = "2" Then asset_type = "2 - Land"
'         If type_code = "3" Then asset_type = "3 - Buildings"
'         If type_code = "4" Then asset_type = "4 - Mobile Home"
'         If type_code = "5" Then asset_type = "5 - Life Estate"
'         If type_code = "6" Then asset_type = "6 - Other"
'
'         If type_verif_code = "TX" Then asset_verification = "TX - Property Tax Statement"
'         If type_verif_code = "PU" Then asset_verification = "PU - Purchase Agreement"
'         If type_verif_code = "TI" Then asset_verification = "TI - Title/Deed"
'         If type_verif_code = "CD" Then asset_verification = "CD - Contract for Deed"
'         If type_verif_code = "CO" Then asset_verification = "CO - County Record"
'         If type_verif_code = "OT" Then asset_verification = "OT - Other Document"
'         If type_verif_code = "NO" Then asset_verification = "NO - No Ver Prvd"
'
'         rest_market_value = replace(rest_market_value, "_", "")
'         rest_market_value = trim(rest_market_value)
'
'         If value_verif_code = "TX" Then rest_value_verification = "TX - Property Tax Statement"
'         If value_verif_code = "PU" Then rest_value_verification = "PU - Purchase Agreement"
'         If value_verif_code = "AP" Then rest_value_verification = "AP - Appraisal"
'         If value_verif_code = "CO" Then rest_value_verification = "CO - County Record"
'         If value_verif_code = "OT" Then rest_value_verification = "OT - Other Document"
'         If value_verif_code = "NO" Then rest_value_verification = "NO - No Ver Prvd"
'
'         rest_amount_owed = replace(rest_amt_owed, "_", "")
'         rest_amount_owed = trim(rest_amount_owed)
'
'         If amt_owed_verif_code = "MO" Then rest_owed_verification = "TI - Title/Deed"
'         If amt_owed_verif_code = "LN" Then rest_owed_verification = "CD - Contract for Deed"
'         If amt_owed_verif_code = "CD" Then rest_owed_verification = "CD - Contract for Deed"
'         If amt_owed_verif_code = "OT" Then rest_owed_verification = "OT - Other Document"
'         If amt_owed_verif_code = "NO" Then rest_owed_verification = "NO - No Ver Prvd"
'
'         rest_owed_date = replace(rest_eff_date, " ", "/")
'         If rest_owed_date = "__/__/__" Then rest_owed_date = ""
'
'         If rest_status = "1" Then rest_property_status = "1 - Home Residence"
'         If rest_status = "2" Then rest_property_status = "2 - For Sale, IV-E Rpymt Agmt"
'         If rest_status = "3" Then rest_property_status = "3 - Joint Owner, Unavailable"
'         If rest_status = "4" Then rest_property_status = "4 - Income Producing"
'         If rest_status = "5" Then rest_property_status = "5 - Future Residence"
'         If rest_status = "6" Then rest_property_status = "6 - Other"
'         If rest_status = "7" Then rest_property_status = "7 - For Sale, Unavailable"
'
'         joint_owners_yn = replace(rest_joint_yn, "_", "")
'         share_ratio = replace(rest_ratio, "_", "")
'
'         rest_ive_repayment_agreement_date = replace(repymt_agree_date, " ", "/")
'         If rest_ive_repayment_agreement_date = "__/__/__" Then rest_ive_repayment_agreement_date = ""
'
'         asset_update_date = replace(rest_update_date, " ", "/")
'
' 	end sub
'
' end class


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

function access_AREP_panel(access_type, arep_name, arep_addr_street, arep_addr_city, arep_addr_state, arep_addr_zip, arep_phone_one, arep_ext_one, arep_phone_two, arep_ext_two, forms_to_arep, mmis_mail_to_arep)

	Call navigate_to_MAXIS_screen("STAT", "AREP")

	EMReadScreen arep_name, 37, 4, 32
	arep_name = replace(arep_name, "_", "")
	If arep_name <> "" Then
		EMReadScreen arep_street_one, 22, 5, 32
		EMReadScreen arep_street_two, 22, 6, 32
		EMReadScreen arep_addr_city, 15, 7, 32
		EMReadScreen arep_addr_state, 2, 7, 55
		EMReadScreen arep_addr_zip, 5, 7, 64

		arep_street_one = replace(arep_street_one, "_", "")
		arep_street_two = replace(arep_street_two, "_", "")
		arep_addr_street = arep_street_one & " " & arep_street_two
		arep_addr_street = trim( arep_addr_street)
		arep_addr_city = replace(arep_addr_city, "_", "")
		arep_addr_state = replace(arep_addr_state, "_", "")
		arep_addr_zip = replace(arep_addr_zip, "_", "")

		state_array = split(state_list, chr(9))
		For each state_item in state_array
			If arep_addr_state = left(state_item, 2) Then
				arep_addr_state = state_item
			End If
		Next

		EMReadScreen arep_phone_one, 14, 8, 34
		EMReadScreen arep_ext_one, 3, 8, 55
		EMReadScreen arep_phone_two, 14, 9, 34
		EMReadScreen arep_ext_two, 3, 8, 55

		arep_phone_one = replace(arep_phone_one, ")", "")
		arep_phone_one = replace(arep_phone_one, "  ", "-")
		arep_phone_one = replace(arep_phone_one, " ", "-")
		If arep_phone_one = "___-___-____" Then arep_phone_one = ""

		arep_phone_two = replace(arep_phone_two, ")", "")
		arep_phone_two = replace(arep_phone_two, "  ", "-")
		arep_phone_two = replace(arep_phone_two, " ", "-")
		If arep_phone_two = "___-___-____" Then arep_phone_two = ""

		arep_ext_one = replace(arep_ext_one, "_", "")
		arep_ext_two = replace(arep_ext_two, "_", "")

		EMReadScreen forms_to_arep, 1, 10, 45
		EMReadScreen mmis_mail_to_arep, 1, 10, 77

	End If

end function

function add_new_HH_MEMB()



end function
' show_pg_one_memb01_and_exp
' show_pg_one_address
' show_pg_memb_list
' show_q_1_6
' show_q_7_11
' show_q_14_15
' show_q_21_24
' show_qual
' show_pg_last
'
' update_addr
' update_pers

function check_for_errors()
	' page_display = show_pg_one_memb01_and_exp
	If who_are_we_completing_the_interview_with = "Select or Type" Then err_msg = err_msg & vbNewLine & "* PICK A PERSON"
	' If  =  Then err_msg = err_msg & vbNewLine & "* "
	' If  =  Then err_msg = err_msg & vbNewLine & "* "
	' If  =  Then err_msg = err_msg & vbNewLine & "* "
	' If  =  Then err_msg = err_msg & vbNewLine & "* "
	' If  =  Then err_msg = err_msg & vbNewLine & "* "
	' If  =  Then err_msg = err_msg & vbNewLine & "* "
	' If  =  Then err_msg = err_msg & vbNewLine & "* "


end function

function define_main_dialog()

	BeginDialog Dialog1, 0, 0, 555, 385, "Full Interview Questions"

	  ButtonGroup ButtonPressed
	    If page_display = show_pg_one_memb01_and_exp Then
			Text 487, 12, 60, 13, "Applicant and EXP"

			ComboBox 120, 10, 205, 45, all_the_clients+chr(9)+who_are_we_completing_the_interview_with, who_are_we_completing_the_interview_with
		    EditBox 305, 45, 50, 15, exp_q_1_income_this_month
		    EditBox 305, 65, 50, 15, exp_q_2_assets_this_month
		    EditBox 305, 85, 50, 15, exp_q_3_rent_this_month
		    CheckBox 120, 105, 30, 10, "Heat", caf_exp_pay_heat_checkbox
		    CheckBox 155, 105, 65, 10, "Air Conditioning", caf_exp_pay_ac_checkbox
		    CheckBox 225, 105, 45, 10, "Electricity", caf_exp_pay_electricity_checkbox
		    CheckBox 275, 105, 35, 10, "Phone", caf_exp_pay_phone_checkbox
		    CheckBox 320, 105, 35, 10, "None", caf_exp_pay_none_checkbox
		    DropListBox 240, 120, 40, 45, "No"+chr(9)+"Yes", exp_migrant_seasonal_formworker_yn
		    DropListBox 360, 135, 40, 45, "No"+chr(9)+"Yes", exp_received_previous_assistance_yn
		    EditBox 75, 155, 80, 15, exp_previous_assistance_when
		    EditBox 195, 155, 85, 15, exp_previous_assistance_where
		    EditBox 315, 155, 85, 15, exp_previous_assistance_what
		    DropListBox 155, 175, 40, 45, "No"+chr(9)+"Yes", exp_pregnant_yn
		    ComboBox 250, 175, 150, 45, all_the_clients, exp_pregnant_who
		    Text 10, 15, 110, 10, "Who are you interviewing with?"
		    GroupBox 5, 35, 400, 160, "Expedited Questions -CAF Answers"
		    Text 15, 50, 270, 10, "1. How much income (cash or checkes) did or will your household get this month?"
		    Text 15, 70, 290, 10, "2. How much does your household (including children) have cash, checking or savings?"
		    Text 15, 90, 225, 10, "3. How much does your household pay for rent/mortgage per month?"
		    Text 25, 105, 90, 10, "What utilities do you pay?"
		    Text 15, 125, 225, 10, "4. Is anyone in your household a migrant or seasonal farm worker?"
		    Text 15, 140, 345, 10, "5. Has anyone in your household ever received cash assistance, commodities or SNAP benefits before?"
		    Text 25, 160, 50, 10, "If yes, When?"
		    Text 165, 160, 30, 10, "Where?"
		    Text 290, 160, 25, 10, "What?"
		    Text 15, 180, 135, 10, "6. Is anyone in your household pregnant?"
		    Text 205, 180, 45, 10, "If yes, who?"
			GroupBox 5, 200, 475, 160, "Expedited Determination"
		    Text 15, 210, 190, 10, "Confirm the Income received in the application month. "
		    Text 20, 220, 230, 10, "What is the total of the income recevied in the month of application?"
		    EditBox 250, 215, 55, 15, intv_app_month_income
		    PushButton 320, 215, 145, 15, "Client is unsure of App Month Income", exp_income_guidance_btn
		    Text 15, 240, 115, 10, "Confirm the Assets the client has."
		    Text 20, 250, 245, 10, "Use the best detail of assets the client has available. Liquid Asset amount?"
		    EditBox 270, 245, 50, 15, intv_app_month_asset
		    Text 15, 270, 195, 10, "Confirm Expenses the client has in the application month."
		    Text 20, 280, 180, 10, "What is the housing expense? (Rent, Mortgage, ectc.)"
		    EditBox 210, 275, 50, 15, intv_app_month_housing_expense
		    Text 20, 295, 115, 10, "What utilities expenses exist?"
		    CheckBox 130, 295, 30, 10, "Heat", intv_exp_pay_heat_checkbox
		    CheckBox 165, 295, 65, 10, "Air Conditioning", intv_exp_pay_ac_checkbox
		    CheckBox 235, 295, 45, 10, "Electricity", intv_exp_pay_electricity_checkbox
		    CheckBox 285, 295, 35, 10, "Phone", intv_exp_pay_phone_checkbox
		    CheckBox 330, 295, 35, 10, "None", intv_exp_pay_none_checkbox
		    Text 15, 315, 105, 10, "Do we have an ID verification?"
		    DropListBox 125, 310, 45, 45, "No"+chr(9)+"Yes", id_verif_on_file
		    Text 195, 315, 165, 10, "Check ECF, SOL-Q, and check in with the client."
		    Text 15, 330, 240, 10, "Is the household active SNAP in another state for the application month?"
		    DropListBox 255, 325, 45, 45, "No"+chr(9)+"Yes", snap_active_in_other_state
		    Text 15, 345, 270, 10, "Was the last SNAP benefit for this case 'Expedited' with postponed verifications?"
		    DropListBox 285, 340, 45, 45, "No"+chr(9)+"Yes", last_snap_was_exp
		ElseIf page_display = show_pg_one_address Then
			Text 500, 27, 60, 13, "CAF ADDR"
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
				DropListBox 205, 70, 110, 45, ""+chr(9)+state_list, resi_addr_state
				EditBox 340, 70, 35, 15, resi_addr_zip
				DropListBox 125, 90, 45, 45, "No"+chr(9)+"Yes", reservation_yn
				EditBox 245, 90, 130, 15, reservation_name
				DropListBox 125, 110, 45, 45, "No"+chr(9)+"Yes", homeless_yn
				DropListBox 245, 110, 130, 45, "Select"+chr(9)+"01 - Own home, lease or roommate"+chr(9)+"02 - Family/Friends - economic hardship"+chr(9)+"03 -  servc prvdr- foster/group home"+chr(9)+"04 - Hospital/Treatment/Detox/Nursing Home"+chr(9)+"05 - Jail/Prison//Juvenile Det."+chr(9)+"06 - Hotel/Motel"+chr(9)+"07 - Emergency Shelter"+chr(9)+"08 - Place not meant for Housing"+chr(9)+"09 - Declined"+chr(9)+"10 - Unknown"+chr(9)+"Blank", living_situation
				EditBox 70, 160, 305, 15, mail_addr_street_full
				EditBox 70, 180, 105, 15, mail_addr_city
				DropListBox 205, 180, 110, 45, ""+chr(9)+state_list, mail_addr_state
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
			Text 10, 5, 400, 10, "Review information for ALL household members, ensuring the information is accurate."
			Text 10, 15, 400, 10, "You must click on each Person button below and on the left to view each person."

			If update_pers = FALSE Then
				Text 70, 45, 90, 15, HH_MEMB_ARRAY(last_name_const, selected_memb)
				Text 165, 45, 75, 15, HH_MEMB_ARRAY(first_name_const, selected_memb)
				Text 245, 45, 50, 15, HH_MEMB_ARRAY(mid_initial, selected_memb)
				Text 300, 45, 175, 15, HH_MEMB_ARRAY(other_names, selected_memb)
				If HH_MEMB_ARRAY(ssn_verif, selected_memb) = "V - System Verified" Then
					Text 70, 75, 70, 15, HH_MEMB_ARRAY(ssn, selected_memb)
				Else
					EditBox 70, 75, 70, 15, HH_MEMB_ARRAY(ssn, selected_memb)
				End If
				Text 145, 75, 70, 15, HH_MEMB_ARRAY(date_of_birth, selected_memb)
				Text 220, 75, 50, 45, HH_MEMB_ARRAY(gender, selected_memb)
				Text 275, 75, 90, 45, HH_MEMB_ARRAY(rel_to_applcnt, selected_memb)
				Text 370, 75, 105, 45, HH_MEMB_ARRAY(marital_status, selected_memb)
				Text 70, 105, 110, 15, HH_MEMB_ARRAY(last_grade_completed, selected_memb)
				Text 195, 105, 70, 15, HH_MEMB_ARRAY(mn_entry_date, selected_memb)
				Text 270, 105, 135, 15, HH_MEMB_ARRAY(former_state, selected_memb)
				Text 400, 105, 75, 45, HH_MEMB_ARRAY(citizen, selected_memb)
				Text 70, 135, 60, 45, HH_MEMB_ARRAY(interpreter, selected_memb)
				Text 140, 135, 120, 15, HH_MEMB_ARRAY(spoken_lang, selected_memb)
				Text 140, 165, 120, 15, HH_MEMB_ARRAY(written_lang, selected_memb)
				Text 330, 145, 40, 45, HH_MEMB_ARRAY(ethnicity_yn, selected_memb)
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
				EditBox 70, 45, 90, 15, HH_MEMB_ARRAY(last_name_const, selected_memb)
				EditBox 165, 45, 75, 15, HH_MEMB_ARRAY(first_name_const, selected_memb)
				EditBox 245, 45, 50, 15, HH_MEMB_ARRAY(mid_initial, selected_memb)
				EditBox 300, 45, 175, 15, HH_MEMB_ARRAY(other_names, selected_memb)
				EditBox 70, 75, 70, 15, HH_MEMB_ARRAY(ssn, selected_memb)
				EditBox 145, 75, 70, 15, HH_MEMB_ARRAY(date_of_birth, selected_memb)
				DropListBox 220, 75, 50, 45, ""+chr(9)+"Male"+chr(9)+"Female", HH_MEMB_ARRAY(gender, selected_memb)
				DropListBox 275, 75, 90, 45, memb_panel_relationship_list, HH_MEMB_ARRAY(rel_to_applcnt, selected_memb)
				DropListBox 370, 75, 105, 45, marital_status_list, HH_MEMB_ARRAY(marital_status, selected_memb)
				EditBox 70, 105, 110, 15, HH_MEMB_ARRAY(last_grade_completed, selected_memb)
				EditBox 185, 105, 70, 15, HH_MEMB_ARRAY(mn_entry_date, selected_memb)
				EditBox 260, 105, 135, 15, HH_MEMB_ARRAY(former_state, selected_memb)
				DropListBox 400, 105, 75, 45, ""+chr(9)+"Yes"+chr(9)+"No", HH_MEMB_ARRAY(citizen, selected_memb)
				DropListBox 70, 135, 60, 45, ""+chr(9)+"Yes"+chr(9)+"No", HH_MEMB_ARRAY(interpreter, selected_memb)
				EditBox 140, 135, 120, 15, HH_MEMB_ARRAY(spoken_lang, selected_memb)
				EditBox 140, 165, 120, 15, HH_MEMB_ARRAY(written_lang, selected_memb)
				DropListBox 330, 145, 40, 45, ""+chr(9)+"Yes"+chr(9)+"No", HH_MEMB_ARRAY(ethnicity_yn, selected_memb)

				PushButton 385, 330, 95, 15, "Save Information", save_information_btn
			End If
			CheckBox 330, 170, 30, 10, "Asian", HH_MEMB_ARRAY(race_a_checkbox, selected_memb)
			CheckBox 330, 180, 30, 10, "Black", HH_MEMB_ARRAY(race_b_checkbox, selected_memb)
			CheckBox 330, 190, 120, 10, "American Indian or Alaska Native", HH_MEMB_ARRAY(race_n_checkbox, selected_memb)
			CheckBox 330, 200, 130, 10, "Pacific Islander and Native Hawaiian", HH_MEMB_ARRAY(race_p_checkbox, selected_memb)
			CheckBox 330, 210, 130, 10, "White", HH_MEMB_ARRAY(race_w_checkbox, selected_memb)
			CheckBox 70, 210, 50, 10, "SNAP (food)", HH_MEMB_ARRAY(snap_req_checkbox, selected_memb)
			CheckBox 125, 210, 65, 10, "Cash programs", HH_MEMB_ARRAY(cash_req_checkbox, selected_memb)
			CheckBox 195, 210, 85, 10, "Emergency Assistance", HH_MEMB_ARRAY(emer_req_checkbox, selected_memb)
			CheckBox 280, 210, 30, 10, "NONE", HH_MEMB_ARRAY(none_req_checkbox, selected_memb)
			DropListBox 70, 250, 80, 45, ""+chr(9)+"Yes"+chr(9)+"No"+chr(9)+HH_MEMB_ARRAY(intend_to_reside_in_mn, selected_memb), HH_MEMB_ARRAY(intend_to_reside_in_mn, selected_memb)
			EditBox 155, 250, 205, 15, HH_MEMB_ARRAY(imig_status, selected_memb)
			DropListBox 365, 250, 55, 45, ""+chr(9)+"Yes"+chr(9)+"No", HH_MEMB_ARRAY(clt_has_sponsor, selected_memb)
			DropListBox 70, 280, 80, 50, "Not Needed"+chr(9)+"Requested"+chr(9)+"On File", HH_MEMB_ARRAY(client_verification, selected_memb)
			EditBox 155, 280, 320, 15, HH_MEMB_ARRAY(client_verification_details, selected_memb)
			EditBox 70, 310, 405, 15, HH_MEMB_ARRAY(client_notes, selected_memb)
			If HH_MEMB_ARRAY(ref_number, selected_memb) = "" Then
				GroupBox 65, 25, 415, 200, "Person " & known_membs+1
				GroupBox 65, 230, 415, 100, "Person " & known_membs+1 & " Interview Questions"
			Else
				GroupBox 65, 25, 415, 200, "Person " & known_membs+1 & " - MEMBER " & HH_MEMB_ARRAY(ref_number, selected_memb)
				GroupBox 65, 230, 415, 100, "Person " & known_membs+1 & " - MEMBER " & HH_MEMB_ARRAY(ref_number, selected_memb) & " Interview Questions"

			End If
			y_pos = 35
			For the_memb = 0 to UBound(HH_MEMB_ARRAY, 2)
				If the_memb = selected_memb Then
					Text 20, y_pos + 1, 45, 10, "Person " & (the_memb + 1)
				Else
					PushButton 10, y_pos, 45, 10, "Person " & (the_memb + 1), HH_MEMB_ARRAY(button_one, the_memb)
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
		ElseIf page_display = show_q_1_6 Then
			Text 510, 57, 60, 13, "Q. 1 - 6"
			y_pos = 10

			GroupBox 5, y_pos, 475, 55, "1. Does everyone in your household buy, fix or eat food with you?"
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_1_yn
			Text 95, y_pos, 25, 10, "write-in:"
			If question_1_verif_yn = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_1_notes
			Else
				EditBox 120, y_pos - 5, 235, 15, question_1_notes
				Text 360, y_pos, 110, 10, "Q1 - Verification - " & question_1_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_1_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_1_btn
			y_pos = y_pos + 20

			GroupBox 5, y_pos, 475, 55, "2. Is anyone in the household, who is age 60 or over or disabled, unable to buy or fix food due to a disability?"
			y_pos = y_pos + 20
			' Text 20, 55, 115, 10, "buy or fix food due to a disability?"
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_2_yn
			Text 95, y_pos, 25, 10, "write-in:"
			If question_2_verif_yn = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_2_notes
			Else
				EditBox 120, y_pos - 5, 235, 15, question_2_notes
				Text 360, y_pos, 110, 10, "Q2 - Verification - " & question_2_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_2_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_2_btn
			y_pos = y_pos + 20

			GroupBox 5, y_pos, 475, 55, "3. Is anyone in the household attending school?"
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_3_yn
			Text 95, y_pos, 25, 10, "write-in:"
			If question_3_verif_yn = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_3_notes
			Else
				EditBox 120, y_pos - 5, 235, 15, question_3_notes
				Text 360, y_pos, 110, 10, "Q3 - Verification - " & question_3_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_3_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_3_btn
			y_pos = y_pos + 20

			GroupBox 5, y_pos, 475, 55, "4. Is anyone in your household temporarily not living in your home? (eg. vacation, foster care, treatment, hospital, job search)"
			y_pos = y_pos + 20
			' Text 20, 135, 230, 10, "(for example: vacation, foster care, treatment, hospital, job search)"
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_4_yn
			Text 95, y_pos, 25, 10, "write-in:"
			If question_4_verif_yn = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_4_notes
			Else
				EditBox 120, y_pos - 5, 235, 15, question_4_notes
				Text 360, y_pos, 110, 10, "Q4 - Verification - " & question_4_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_4_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_4_btn
			y_pos = y_pos + 20

			GroupBox 5, y_pos, 475, 55, "5. Is anyone blind, or does anyone have a physical or mental health condition that limits the ability to work or perform daily activities?"
			y_pos = y_pos + 20
			' Text 20, 180, 185, 10, " that limits the ability to work or perform daily activities?"
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_5_yn
			Text 95, y_pos, 25, 10, "write-in:"
			If question_5_verif_yn = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_5_notes
			Else
				EditBox 120, y_pos - 5, 235, 15, question_5_notes
				Text 360, y_pos, 110, 10, "Q5 - Verification - " & question_5_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_5_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_5_btn
			y_pos = y_pos + 20

			GroupBox 5, y_pos, 475, 55, "6. Is anyone unable to work for reasons other than illness or disability?"
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_6_yn
			Text 95, y_pos, 25, 10, "write-in:"
			If question_6_verif_yn = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_6_notes
			Else
				EditBox 120, y_pos - 5, 235, 15, question_6_notes
				Text 360, y_pos, 110, 10, "Q6 - Verification - " & question_6_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_6_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_6_btn
			y_pos = y_pos + 20

		ElseIf page_display = show_q_7_11 Then
			Text 508, 72, 60, 13, "Q. 7 - 11"
			y_pos = 10

			GroupBox 5, y_pos, 475, 55, "7. In the last 60 days did anyone in the household: - Stop working or quit a job? - Refuse a job offer? - Ask to work fewer hours? - Go on strike?"
			' Text 20, 315, 350, 10, "- Stop working or quit a job?   - Refuse a job offer? - Ask to work fewer hours?   - Go on strike?"
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_7_yn
			Text 95, y_pos, 25, 10, "write-in:"
			If question_7_verif_yn = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_7_notes
			Else
				EditBox 120, y_pos - 5, 235, 15, question_7_notes
				Text 360, y_pos, 110, 10, "Q7 - Verification - " & question_7_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_7_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_7_btn
			y_pos = y_pos + 20

			GroupBox 5, y_pos, 475, 65, "8. Has anyone in the household had a job or been self-employed in the past 12 months?"
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_8_yn
			Text 95, y_pos, 25, 10, "write-in:"
			If question_8_verif_yn = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_8_notes
			Else
				EditBox 120, y_pos - 5, 235, 15, question_8_notes
				Text 360, y_pos, 110, 10, "Q8 - Verification - " & question_8_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 400, 10, "a. FOR SNAP ONLY: Has anyone in the household had a job or been self-employed in the past 36 months?       CAF Answer"
			DropListBox 415, y_pos - 5, 35, 45, question_answers, question_8a_yn
			y_pos = y_pos + 15
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_8_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_8_btn
			y_pos = y_pos + 25

			grp_len = 35
			for each_job = 0 to UBOUND(JOBS_ARRAY, 2)
				' If JOBS_ARRAY(jobs_employer_name, each_job) <> "" AND JOBS_ARRAY(jobs_employee_name, each_job) <> "" AND JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" AND JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then
				If JOBS_ARRAY(jobs_employer_name, each_job) <> "" OR JOBS_ARRAY(jobs_employee_name, each_job) <> "" OR JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" OR JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then grp_len = grp_len + 20
			next
			GroupBox 5, y_pos, 475, grp_len, "9. Does anyone in the household have a job or expect to get income from a job this month or next month?"
			PushButton 425, y_pos, 55, 10, "ADD JOB", add_job_btn
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_9_yn
			Text 95, y_pos, 25, 10, "write-in:"
			EditBox 120, y_pos - 5, 350, 15, question_9_notes
			' Text 360, y_pos, 110, 10, "Q9 - Verification - " & question_9_verif_yn
			' y_pos = y_pos + 20

			' PushButton 300, 100, 75, 10, "ADD VERIFICATION", add_verif_9_btn
			' y_pos = 110
			' If JOBS_ARRAY(jobs_employee_name, 0) <> "" Then
			First_job = TRUE
			for each_job = 0 to UBOUND(JOBS_ARRAY, 2)
				' If JOBS_ARRAY(jobs_employer_name, each_job) <> "" AND JOBS_ARRAY(jobs_employee_name, each_job) <> "" AND JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" AND JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then
				If JOBS_ARRAY(jobs_employer_name, each_job) <> "" OR JOBS_ARRAY(jobs_employee_name, each_job) <> "" OR JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" OR JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then
					If First_job = TRUE Then y_pos = y_pos + 20
					First_job = FALSE
					If JOBS_ARRAY(verif_yn, each_job) = "" Then Text 15, y_pos, 395, 10, "Employer: " & JOBS_ARRAY(jobs_employer_name, each_job) & "  - Employee: " & JOBS_ARRAY(jobs_employee_name, each_job) & "   - Gross Monthly Earnings: $ " & JOBS_ARRAY(jobs_gross_monthly_earnings, each_job)
					If JOBS_ARRAY(verif_yn, each_job) <> "" Then Text 15, y_pos, 395, 10, "Employer: " & JOBS_ARRAY(jobs_employer_name, each_job) & "  - Employee: " & JOBS_ARRAY(jobs_employee_name, each_job) & "   - Gross Monthly Earnings: $ " & JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) & "   - Verification - " & JOBS_ARRAY(verif_yn, each_job)
					PushButton 450, y_pos, 20, 10, "EDIT", JOBS_ARRAY(jobs_edit_btn, each_job)
					y_pos = y_pos + 10
				End If
			next
			If First_job = TRUE Then y_pos = y_pos + 10
			y_pos = y_pos + 15

			GroupBox 5, y_pos, 475, 55, "10. Is anyone in the household self-employed or does anyone expect to get income from self-employment this month or next month?"
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_10_yn
			Text 95, y_pos, 50, 10, "Gross Earnings:"
			EditBox 145, y_pos - 5, 35, 15, question_10_monthly_earnings
			Text 180, y_pos, 25, 10, "write-in:"
			If question_10_verif_yn = "" Then
				EditBox 205, y_pos - 5, 270, 15, question_10_notes
			Else
				EditBox 205, y_pos - 5, 150, 15, question_10_notes
				Text 360, y_pos, 105, 10, "Q10 - Verification - " & question_10_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_10_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_10_btn
			y_pos = y_pos + 20

			GroupBox 5, y_pos, 475, 55, "11. Do you expect any changes in income, expenses or work hours?"
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_11_yn
			Text 95, y_pos, 25, 10, "write-in:"
			If question_11_verif_yn = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_11_notes
			Else
				EditBox 120, y_pos - 5, 235, 15, question_11_notes
				Text 360, y_pos, 110, 10, "Q11 - Verification - " & question_11_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_11_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_11_btn
			y_pos = y_pos + 25

			Text 5, y_pos, 75, 10, "Pricipal Wage Earner"
			DropListBox 85, y_pos - 5, 175, 45, all_the_clients, pwe_selection
			y_pos = y_pos + 10


		ElseIf page_display = show_q_12_13 Then
			Text 505, 87, 60, 13, "Q. 12 - 13"
			y_pos = 10

			GroupBox 5, y_pos, 475, 125, "12. Has anyone in the household applied for or does anyone get any of the following type of income each month?"
			' y_pos = y_pos + 15

			y_pos = y_pos + 20
			col_1_1 = 15
			col_1_2 = 55
			col_1_3 = 115

			col_2_1 = 165
			col_2_2 = 205
			col_2_3 = 260

			col_3_1 = 320
			col_3_2 = 360
			col_3_3 = 430

			Text 	col_1_1, 		y_pos, 40, 10, "CAF Answer"
			Text 	col_1_3 - 3, 	y_pos, 40, 10, "CAF Amount"
			Text 	col_2_1, 		y_pos, 40, 10, "CAF Answer"
			Text 	col_2_3 - 3, 	y_pos, 40, 10, "CAF Amount"
			Text 	col_3_1, 		y_pos, 40, 10, "CAF Answer"
			Text 	col_3_3 - 3, 	y_pos, 40, 10, "CAF Amount"
			y_pos = y_pos + 15

			DropListBox 	col_1_1, 	y_pos, 		35, 45, question_answers, question_12_rsdi_yn
			Text 			col_1_2, 	y_pos + 5, 	60, 10, "RSDI                  $"
			EditBox 		col_1_3,	y_pos, 		35, 15, question_12_rsdi_amt
			DropListBox 	col_2_1, 	y_pos, 		35, 45, question_answers, question_12_ssi_yn
			Text 			col_2_2, 	y_pos + 5, 	60, 10, "SSI                $"
			EditBox 		col_2_3, 	y_pos, 		35, 15, question_12_ssi_amt
			DropListBox 	col_3_1, 	y_pos, 		35, 45, question_answers, question_12_va_yn
			Text 			col_3_2, 	y_pos + 5, 	70, 10, "VA                          $"
			EditBox 		col_3_3, 	y_pos, 		35, 15, question_12_va_amt
			y_pos = y_pos + 15

			DropListBox 	col_1_1, 	y_pos, 		35, 45, question_answers, question_12_ui_yn
			Text 			col_1_2, 	y_pos + 5, 	60, 10, "UI                       $"
			EditBox 		col_1_3, 	y_pos, 		35, 15, question_12_ui_amt
			DropListBox 	col_2_1, 	y_pos, 		35, 45, question_answers, question_12_wc_yn
			Text 			col_2_2, 	y_pos + 5, 	60, 10, "WC                $"
			EditBox 		col_2_3, 	y_pos, 		35, 15, question_12_wc_amt
			DropListBox 	col_3_1, 	y_pos, 		35, 45, question_answers, question_12_ret_yn
			Text 			col_3_2, 	y_pos + 5, 	85, 10, "Retirement Ben.     $"
			EditBox 		col_3_3, 	y_pos, 		35, 15, question_12_ret_amt
			y_pos = y_pos + 15

			DropListBox 	col_1_1, 	y_pos, 		35, 45, question_answers, question_12_trib_yn
			Text 			col_1_2, 	y_pos + 5, 	60, 10, "Tribal Payments  $"
			EditBox 		col_1_3, 	y_pos, 		35, 15, question_12_trib_amt
			DropListBox 	col_2_1, 	y_pos, 		35, 45, question_answers, question_12_cs_yn
			Text 			col_2_2, 	y_pos + 5, 	60, 10, "CSES             $"
			EditBox 		col_2_3,	y_pos, 		35, 15, question_12_cs_amt
			DropListBox 	col_3_1, 	y_pos, 		35, 45, question_answers, question_12_other_yn
			Text 			col_3_2, 	y_pos + 5, 	110, 10, "Other unearned       $"
			EditBox 		col_3_3, 	y_pos, 		35, 15, question_12_other_amt
			y_pos = y_pos + 25

			Text 15, y_pos, 25, 10, "Write-in:"
			If question_12_verif_yn = "" Then
				EditBox 40, y_pos - 5, 435, 15, question_12_notes
			Else
				EditBox 40, y_pos - 5, 315, 15, question_12_notes
				Text 360, y_pos, 110, 10, "Q12 - Verification - " & question_12_verif_yn
			End If
			' Text 360, y_pos, 105, 10, "Q10 - Verification - " & question_10_verif_yn
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_11_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_12_btn
			y_pos = y_pos + 25

			' Text 0, 0, 0, 0, ""
			' Text 0, 0, 0, 0, ""
			' Text 0, 0, 0, 0, ""
			' Text 0, 0, 0, 0, ""
			' Text 0, 0, 0, 0, ""
			GroupBox 5, y_pos, 475, 55, "13. Does anyone in the household have or expect to get any loans, scholarships or grants for attending school?"
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_13_yn
			Text 95, y_pos, 25, 10, "write-in:"
			If question_13_verif_yn = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_13_notes
			Else
				EditBox 120, y_pos - 5, 235, 15, question_13_notes
				Text 360, y_pos, 110, 10, "Q13 - Verification - " & question_13_verif_yn
			End If
			y_pos = y_pos + 20

			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_13_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_13_btn
			y_pos = y_pos + 20

		ElseIf page_display = show_q_14_15 Then
			Text 505, 102, 60, 13, "Q. 14 - 15"
			y_pos = 10

			GroupBox 5, 10, 475, 130, "14. Does your household have the following housing expenses?"
			y_pos = y_pos + 15
			col_1_1 = 15
			col_1_2 = 85
			col_2_1 = 220
			col_2_2 = 290

			Text 	col_1_1, 		y_pos, 40, 10, "CAF Answer"
			Text 	col_2_1, 		y_pos, 40, 10, "CAF Answer"
			y_pos = y_pos + 15

			DropListBox 	col_1_1, y_pos - 5, 60, 45, question_answers, question_14_rent_yn
			Text 			col_1_2, y_pos, 	70, 10, "Rent"
			DropListBox 	col_2_1, y_pos - 5, 60, 45, question_answers, question_14_subsidy_yn
			Text 			col_2_2, y_pos, 	100, 10, "Rent or Section 8 Subsidy"
			y_pos = y_pos + 15

			DropListBox 	col_1_1, y_pos - 5, 60, 45, question_answers, question_14_mortgage_yn
			Text 			col_1_2, y_pos, 	125, 10, "Mortgage/contract for deed payment"
			DropListBox 	col_2_1, y_pos - 5, 60, 45, question_answers, question_14_association_yn
			Text 			col_2_2, y_pos, 	70, 10, "Association fees"
			y_pos = y_pos + 15

			DropListBox 	col_1_1, y_pos - 5, 60, 45, question_answers, question_14_insurance_yn
			Text 			col_1_2, y_pos, 	85, 10, "Homeowner's insurance"
			DropListBox 	col_2_1, y_pos - 5, 60, 45, question_answers, question_14_room_yn
			Text 			col_2_2, y_pos, 	70, 10, "Room and/or board"
			y_pos = y_pos + 15

			DropListBox 	col_1_1, y_pos - 5, 60, 45, question_answers, question_14_taxes_yn
			Text 			col_1_2, y_pos, 	100, 10, "Real estate taxes"
			y_pos = y_pos + 20

			Text 15, y_pos, 25, 10, "Write-in:"
			If question_14_verif_yn = "" Then
				EditBox 40, y_pos - 5, 435, 15, question_14_notes
			Else
				EditBox 40, y_pos - 5, 315, 15, question_14_notes
				Text 360, y_pos, 110, 10, "Q14 - Verification - " & question_14_verif_yn
			End If
			y_pos = y_pos + 20

			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_14_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_14_btn
			y_pos = y_pos + 25


			GroupBox 5, y_pos, 475, 120, "15. Does your household have the following utility expenses any time during the year? "
			y_pos = y_pos + 15

			col_1_1 = 15
			col_1_2 = 85

			col_2_1 = 185
			col_2_2 = 255

			col_3_1 = 335
			col_3_2 = 405

			Text 	col_1_1, 		y_pos, 40, 10, "CAF Answer"
			Text 	col_2_1, 		y_pos, 40, 10, "CAF Answer"
			Text 	col_3_1, 		y_pos, 40, 10, "CAF Answer"
			y_pos = y_pos + 15

			DropListBox 	col_1_1, y_pos - 5, 35, 45, question_answers, question_15_heat_ac_yn
			Text 			col_1_2, y_pos, 	85, 10, "Heating/air conditioning"
			DropListBox 	col_2_1, y_pos - 5, 35, 45, question_answers, question_15_electricity_yn
			Text 			col_2_2, y_pos, 	70, 10, "Electricity"
			DropListBox 	col_3_1, y_pos - 5, 35, 45, question_answers, question_15_cooking_fuel_yn
			Text 			col_3_2, y_pos, 	70, 10, "Cooking fuel"
			y_pos = y_pos + 15

			DropListBox 	col_1_1, y_pos - 5, 35, 45, question_answers, question_15_water_and_sewer_yn
			Text 			col_1_2, y_pos, 	75, 10, "Water and sewer"
			DropListBox 	col_2_1, y_pos - 5, 35, 45, question_answers, question_15_garbage_yn
			Text 			col_2_2, y_pos, 	60, 10, "Garbage removal"
			DropListBox 	col_3_1, y_pos - 5, 35, 45, question_answers, question_15_phone_yn
			Text 			col_3_2, y_pos, 	70, 10, "Phone/cell phone"
			y_pos = y_pos + 15

			DropListBox 	col_1_1, y_pos - 5, 35, 45, question_answers, question_15_liheap_yn
			Text 			col_1_2, y_pos, 375, 10, "Did you or anyone in your household receive LIHEAP (energy assistance) of more than $20 in the past 12 months?"
			y_pos = y_pos + 20

			Text 15, y_pos, 25, 10, "Write-in:"
			If question_15_verif_yn = "" Then
				EditBox 40, y_pos - 5, 435, 15, question_15_notes
			Else
				EditBox 40, y_pos - 5, 315, 15, question_15_notes
				Text 360, y_pos, 110, 10, "Q15 - Verification - " & question_15_verif_yn
			End If
			y_pos = y_pos + 20

			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_15_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_15_btn

		ElseIf page_display = show_q_16_20 Then
			Text 505, 117, 60, 13, "Q. 16 - 20"
			y_pos = 10

			GroupBox 5, y_pos, 475, 55, "16. Do you or anyone living with you have costs for care of a child(ren) because you or they are working, looking for work or going to school?"
			' Text 95, 200, 125, 10, "looking for work or going to school?"
			y_pos = y_pos + 20
			Text 		15, y_pos, 		40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 	35, 45, question_answers, question_16_yn
			Text 		95, y_pos, 		25, 10, "write-in:"
			If question_16_verif_yn = "" Then
				EditBox 	120, y_pos - 5, 355, 15, question_16_notes
			Else
				EditBox 	120, y_pos - 5, 235, 15, question_16_notes
				Text 		360, y_pos, 	110, 10, "Q16 - Verification - " & question_16_verif_yn
			End If
			y_pos = y_pos + 20
			Text 		15, y_pos, 		60, 10, "Interview Notes:"
			EditBox 	75, y_pos - 5, 	320, 15, question_16_interview_notes
			PushButton 	400, y_pos, 	75, 10, "ADD VERIFICATION", add_verif_16_btn
			y_pos = y_pos + 20

			GroupBox 5, y_pos, 475, 55, "17. Does anyone have costs for care of an ill/disabled adult because you or they are working, looking for work or going to school?"
			' Text 95, 245, 125, 10, "looking for work or going to school?"
			y_pos = y_pos + 20
			Text 		15, y_pos, 		40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 	35, 45, question_answers, question_17_yn
			Text 		95, y_pos, 		25, 10, "write-in:"
			If question_17_verif_yn = "" Then
				EditBox 	120, y_pos - 5, 355, 15, question_17_notes
			Else
				EditBox 	120, y_pos - 5, 235, 15, question_17_notes
				Text 		360, y_pos, 	110, 10, "Q17 - Verification - " & question_17_verif_yn
			End If
			y_pos = y_pos + 20
			Text 		15, y_pos, 		60, 10, "Interview Notes:"
			EditBox 	75, y_pos - 5, 	320, 15, question_17_interview_notes
			PushButton 	400, y_pos, 	75, 10, "ADD VERIFICATION", add_verif_17_btn
			y_pos = y_pos + 20

			GroupBox 5, y_pos, 475, 55, "18. Does anyone in the household pay support, or contribute to a tax dependent who does not live in your home?"
			' Text 95, 290, 215, 10, "or contribute to a tax dependent who does not live in your home?"
			y_pos = y_pos + 20
			Text 		15, y_pos, 		40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 	35, 45, question_answers, question_18_yn
			Text 		95, y_pos, 		25, 10, "write-in:"
			If question_18_verif_yn = "" Then
				EditBox 	120, y_pos - 5, 355, 15, question_18_notes
			Else
				EditBox 	120, y_pos - 5, 235, 15, question_18_notes
				Text 		360, y_pos, 	110, 10, "Q18 - Verification - " & question_18_verif_yn
			End If
			y_pos = y_pos + 20
			Text 		15, y_pos, 		60, 10, "Interview Notes:"
			EditBox 	75, y_pos - 5, 	320, 15, question_18_interview_notes
			PushButton 	400, y_pos, 	75, 10, "ADD VERIFICATION", add_verif_18_btn
			y_pos = y_pos + 20

			' Text 0, 0, 0, 0, ""
			' Text 0, 0, 0, 0, ""
			GroupBox 5, y_pos, 475, 55, "19. For SNAP only: Does anyone in the household have medical expenses? "
			y_pos = y_pos + 20
			Text 		15, y_pos, 		40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 	35, 45, question_answers, question_19_yn
			Text 		95, y_pos, 		25, 10, "write-in:"
			If question_19_verif_yn = "" Then
				EditBox 	120, y_pos - 5, 355, 15, question_19_notes
			Else
				EditBox 	120, y_pos - 5, 235, 15, question_19_notes
				Text 		360, y_pos, 	110, 10, "Q19 - Verification - " & question_19_verif_yn
			End If
			y_pos = y_pos + 20
			Text 		15, y_pos, 60, 	10, "Interview Notes:"
			EditBox 	75, y_pos - 5, 	320, 15, question_19_interview_notes
			PushButton 	400, y_pos, 	75, 10, "ADD VERIFICATION", add_verif_19_btn
			y_pos = y_pos + 20

			GroupBox 5, y_pos, 475, 100, "20. Does anyone in the household own, or is anyone buying, any of the following?"
			y_pos = y_pos + 10
			col_1_1 = 25
			col_1_2 = 90
			col_2_1 = 230
			col_2_2 = 295

			Text 	col_1_1, 		y_pos, 40, 10, "CAF Answer"
			Text 	col_2_1, 		y_pos, 40, 10, "CAF Answer"
			y_pos = y_pos + 15

			DropListBox 	col_1_1, y_pos - 5, 60, 45, question_answers, question_20_cash_yn
			Text 			col_1_2, y_pos, 	70, 10, "Cash"
			DropListBox 	col_2_1, y_pos - 5, 60, 45, question_answers, question_20_acct_yn
			Text 			col_2_2, y_pos, 	175, 10, "Bank accounts (savings, checking, debit card, etc.)"
			y_pos = y_pos + 15

			DropListBox 	col_1_1, y_pos - 5, 60, 45, question_answers, question_20_secu_yn
			Text 			col_1_2, y_pos, 	125, 10, "Stocks, bonds, annuities, 401k, etc."
			DropListBox 	col_2_1, y_pos - 5, 60, 45, question_answers, question_20_cars_yn
			Text 			col_2_2, y_pos, 	180, 10, "Vehicles (cars, trucks, motorcycles, campers, trailers)"
			y_pos = y_pos + 20

			Text 15, y_pos, 25, 10, "Write-in:"
			If question_20_verif_yn = "" Then
				EditBox 40, y_pos - 5, 435, 15, question_20_notes
			Else
				EditBox 40, y_pos - 5, 315, 15, question_20_notes
				Text 360, y_pos, 110, 10, "Q20 - Verification - " & question_20_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_20_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_20_btn
			y_pos = y_pos + 25

		ElseIf page_display = show_q_21_24 Then
			Text 505, 132, 60, 13, "Q. 21 - 24"
			y_pos = 10

			GroupBox 5, y_pos, 475, 55, "21. For Cash programs only: Has anyone in the household given away, sold or traded anything of value in the past 12 months? "
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_21_yn
			Text 95, y_pos, 25, 10, "Write-in:"
			If question_21_verif_yn = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_21_notes
			Else
				EditBox 120, y_pos - 5, 235, 15, question_21_notes
				Text 360, y_pos, 110, 10, "Q21 - Verification - " & question_21_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_21_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_21_btn
			y_pos = y_pos + 25

			GroupBox 5, y_pos, 475, 55, "22. For recertifications only: Did anyone move in or out of your home in the past 12 months?"
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_22_yn
			Text 95, y_pos, 25, 10, "Write-in:"
			If question_22_verif_yn = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_22_notes
			Else
				EditBox 120, y_pos - 5, 235, 15, question_22_notes
				Text 360, y_pos, 110, 10, "Q22 - Verification - " & question_22_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_22_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_22_btn
			y_pos = y_pos + 25

			GroupBox 5, y_pos, 475, 55, "23. For children under the age of 19, are both parents living in the home?"
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_23_yn
			Text 95, y_pos, 25, 10, "Write-in:"
			If question_23_verif_yn = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_23_notes
			Else
	 			EditBox 120, y_pos - 5, 235, 15, question_23_notes
				Text 360, y_pos, 110, 10, "Q23 - Verification - " & question_23_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_23_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_23_btn
			y_pos = y_pos + 25

			GroupBox 5, y_pos, 475, 100, "24. For MSA recipients only: Does anyone in the household have any of the following expenses?"
			y_pos = y_pos + 10

			col_1_1 = 25
			col_1_2 = 90
			col_2_1 = 230
			col_2_2 = 295

			Text 	col_1_1, 		y_pos, 40, 10, "CAF Answer"
			Text 	col_2_1, 		y_pos, 40, 10, "CAF Answer"
			y_pos = y_pos + 15

			DropListBox col_1_1, y_pos - 5, 60, 45, question_answers, question_24_rep_payee_yn
			Text 		col_1_2, y_pos, 	95, 10, "Representative Payee fees"
			DropListBox col_2_1, y_pos - 5, 60, 45, question_answers, question_24_guardian_fees_yn
			Text 		col_2_2, y_pos, 	105, 10, "Guardian Conservator fees"
			y_pos = y_pos + 15

			DropListBox col_1_1, y_pos - 5, 60, 45, question_answers, question_24_special_diet_yn
			Text 		col_1_2, y_pos, 	125, 10, "Physician-perscribed special diet"
			DropListBox col_2_1, y_pos - 5, 60, 45, question_answers, question_24_high_housing_yn
			Text 		col_2_2, y_pos, 	105, 10, "High housing costs"
			y_pos = y_pos + 20

			Text 15, y_pos, 25, 10, "Write-in:"
			If question_24_verif_yn = "" Then
				EditBox 40, y_pos - 5, 435, 15, question_24_notes
			Else
				EditBox 40, y_pos - 5, 315, 15, question_24_notes
				Text 360, y_pos, 110, 10, "Q24 - Verification - " & question_24_verif_yn
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_24_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", add_verif_24_btn

		ElseIf page_display = show_qual Then
			Text 497, 147, 60, 13, "CAF QUAL Q"

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
			Text 495, 162, 60, 13, "CAF Last Page"

			GroupBox 5, 10, 475, 100, "Confirm Authorized Representative"

			If arep_in_MAXIS = TRUE Then Text 15, 25, 120, 10, "Current Authorized Representative"
			If arep_in_MAXIS = FALSE Then Text 15, 25, 120, 10, "There is no Authorized Representative"

			If update_arep = FALSE Then
				Text 85, 55, 170, 15, arep_name
				Text 260, 55, 120, 45, arep_relationship
				Text 385, 55, 85, 15, arep_phone_number
				Text 85, 90, 170, 15, arep_addr_street
				Text 260, 90, 85, 15, arep_addr_city
				Text 350, 90, 65, 45, arep_addr_state
				Text 420, 90, 50, 15, arep_addr_zip
				PushButton 385, 20, 85, 15, "Update AREP Detail", update_information_btn
			End If

			If update_arep = TRUE Then
				EditBox 85, 50, 170, 15, arep_name
				ComboBox 260, 50, 120, 45, "", arep_relationship
				EditBox 385, 50, 85, 15, arep_phone_one
				EditBox 85, 85, 170, 15, arep_addr_street
				EditBox 260, 85, 85, 15, arep_addr_city
				DropListBox 350, 85, 65, 45, state_list, arep_addr_state
				EditBox 420, 85, 50, 15, arep_addr_zip
				PushButton 385, 20, 85, 15, "Save AREP Detail", save_information_btn
			End If

		    Text 85, 40, 45, 10, "AREP Name"
			Text 260, 40, 50, 10, "Relationship"
			Text 385, 40, 50, 10, "Phone Number"
			Text 85, 75, 35, 10, "Address"
			Text 260, 75, 25, 10, "City"
			Text 350, 75, 25, 10, "State"
			Text 420, 75, 35, 10, "Zip Code"
		    CheckBox 10, 50, 55, 10, "Fill out forms", arep_complete_forms_checkbox
		    CheckBox 10, 65, 50, 10, "Get notices", arep_get_notices_checkbox
		    CheckBox 10, 80, 65, 10, "Get and use my", arep_use_SNAP_checkbox
		    Text 20, 90, 50, 10, "SNAP benefits"

		    GroupBox 5, 115, 475, 105, "Signatures"
		    Text 10, 135, 90, 10, "Signature of Primary Adult"
		    ComboBox 105, 130, 110, 45, "Select or Type"+chr(9)+"Signature Completed"+chr(9)+"Blank"+chr(9)+"Accepted Verbally"+chr(9)+"Not Required"+chr(9)+signature_detail, signature_detail
		    Text 220, 135, 25, 10, "person"
		    ComboBox 250, 130, 115, 45, all_the_clients+chr(9)+signature_person, signature_person
		    Text 375, 135, 20, 10, "date"
		    EditBox 400, 130, 50, 15, signature_date
		    Text 10, 155, 90, 10, "Signature of Other Adult"
		    ComboBox 105, 150, 110, 45, "Select or Type"+chr(9)+"Signature Completed"+chr(9)+"Blank"+chr(9)+"Accepted Verbally"+chr(9)+"Not Required"+chr(9)+second_signature_detail, second_signature_detail
		    Text 220, 155, 25, 10, "person"
		    ComboBox 250, 150, 115, 45, all_the_clients+chr(9)+second_signature_person, second_signature_person
		    Text 375, 155, 20, 10, "date"
		    EditBox 400, 150, 50, 15, second_signature_date

			Text 270, 180, 120, 10, "Cient signature accepted verbally?"
			DropListBox 390, 175, 60, 45, "Select..."+chr(9)+"Yes"+chr(9)+"No", client_signed_verbally_yn
			Text 335, 200, 50, 10, "Interview Date:"
			EditBox 390, 195, 60, 15, interview_date

			GroupBox 5, 225, 475, 115, "Benefit Detail"
			' appears_expedited
			' expedited_delay_info
			' expedited_info_does_not_match
			' mismatch_explanation

	    ' ElseIf page_display =  Then
		End If


		If page_display <> show_pg_one_memb01_and_exp 	Then PushButton 485, 10, 65, 13, "Applicant and EXP", caf_page_one_btn
		If page_display <> show_pg_one_address 			Then PushButton 485, 25, 65, 13, "CAF ADDR", caf_addr_btn
		' If page_display <> show_pg_memb_list AND page_display <> show_pg_memb_info AND  page_display <> show_pg_imig Then PushButton 485, 25, 60, 13, "CAF MEMBs", caf_membs_btn
		If page_display <> show_pg_memb_list 			Then PushButton 485, 40, 65, 13, "CAF MEMBs", caf_membs_btn
		If page_display <> show_q_1_6 					Then PushButton 485, 55, 65, 13, "Q. 1 - 6", caf_q_1_6_btn
		If page_display <> show_q_7_11 					Then PushButton 485, 70, 65, 13, "Q. 7 - 11", caf_q_7_11_btn
		If page_display <> show_q_12_13 				Then PushButton 485, 85, 65, 13, "Q. 12 - 13", caf_q_12_13_btn
		If page_display <> show_q_14_15 				Then PushButton 485, 100, 65, 13, "Q. 14 - 15", caf_q_14_15_btn
		If page_display <> show_q_16_20 				Then PushButton 485, 115, 65, 13, "Q. 16 - 20", caf_q_16_20_btn
		If page_display <> show_q_21_24 				Then PushButton 485, 130, 65, 13, "Q. 21 - 24", caf_q_21_24_btn
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
		If page_display <> show_qual 					Then PushButton 485, 145, 65, 13, "CAF QUAL Q", caf_qual_q_btn
		If page_display <> show_pg_last 				Then PushButton 485, 160, 65, 13, "CAF Last Page", caf_last_page_btn
		PushButton 415, 365, 50, 15, "NEXT", next_btn
		PushButton 465, 365, 80, 15, "Complete Interview", finish_interview_btn

	EndDialog

end function

function dialog_movement()
	' case_has_imig = FALSE
	' MsgBox ButtonPressed
	For i = 0 to Ubound(HH_MEMB_ARRAY, 2)
		' If HH_MEMB_ARRAY(i).imig_exists = TRUE Then case_has_imig = TRUE
		' MsgBox HH_MEMB_ARRAY(i).button_one
		If ButtonPressed = HH_MEMB_ARRAY(button_one, i) Then
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

	If ButtonPressed = update_information_btn Then
		If page_display = show_pg_one_address Then update_addr = TRUE
		If page_display = show_pg_memb_list Then update_pers = TRUE
		If page_display = show_pg_last Then update_arep = TRUE
	End If
	If ButtonPressed = save_information_btn Then
		If page_display = show_pg_one_address Then update_addr = FALSE
		If page_display = show_pg_memb_list Then update_pers = FALSE
		If page_display = show_pg_last Then update_arep = FALSE

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
		If memb_selected > UBound(HH_MEMB_ARRAY, 2) Then ButtonPressed = next_btn
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
	If ButtonPressed = add_person_btn Then
		last_clt = UBound(HH_MEMB_ARRAY, 2)
		new_clt = last_clt + 1
		ReDim Preserve HH_MEMB_ARRAY(new_clt)
		Set HH_MEMB_ARRAY(new_clt) = new mx_hh_member
		HH_MEMB_ARRAY(new_clt).button_one = 500 + new_clt
		HH_MEMB_ARRAY(new_clt).button_two = 600 + new_clt
		selected_memb = new_clt
		update_pers = TRUE
	End If
	If ButtonPressed = exp_income_guidance_btn Then
		call guide_through_app_month_income
	End If
	If ButtonPressed = -1 Then ButtonPressed = next_btn
	If ButtonPressed = next_btn Then
		If page_display = show_pg_one_memb01_and_exp 	Then ButtonPressed = caf_addr_btn
		If page_display = show_pg_one_address 			Then ButtonPressed = caf_membs_btn
		If page_display = show_pg_memb_list 			Then ButtonPressed = caf_q_1_6_btn
		If page_display = show_q_1_6 					Then ButtonPressed = caf_q_7_11_btn
		If page_display = show_q_7_11 					Then ButtonPressed = caf_q_12_13_btn
		If page_display = show_q_12_13 					Then ButtonPressed = caf_q_14_15_btn
		If page_display = show_q_14_15 					Then ButtonPressed = caf_q_16_20_btn
		If page_display = show_q_16_20 					Then ButtonPressed = caf_q_21_24_btn
		If page_display = show_q_21_24 					Then ButtonPressed = caf_qual_q_btn
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
	If ButtonPressed = caf_q_1_6_btn Then
		page_display = show_q_1_6
	End If
	If ButtonPressed = caf_q_7_11_btn Then
		page_display = show_q_7_11
	End If
	If ButtonPressed = caf_q_12_13_btn Then
		page_display = show_q_12_13
	End If
	If ButtonPressed = caf_q_14_15_btn Then
		page_display = show_q_14_15
	End If
	If ButtonPressed = caf_q_16_20_btn Then
		page_display = show_q_16_20
	End If
	If ButtonPressed = caf_q_21_24_btn Then
		page_display = show_q_21_24
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

function guide_through_app_month_income()
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 451, 350, "Questions to Guide Determination of Income in Month of Application "
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
	  ButtonGroup ButtonPressed
	    PushButton 395, 330, 50, 15, "Return", return_btn
	EndDialog

	dialog Dialog1

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
		local_changelog_path = user_myDocs_folder & "interview-answers-" & MAXIS_case_number & "-info.txt"
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
			objTextStream.WriteLine "PRE - WHO - " & who_are_we_completing_the_interview_with

			objTextStream.WriteLine "EXP - 1 - " & exp_q_1_income_this_month
			objTextStream.WriteLine "EXP - 2 - " & exp_q_2_assets_this_month
			objTextStream.WriteLine "EXP - 3 - RENT - " & exp_q_3_rent_this_month
			If caf_exp_pay_heat_checkbox = checked 			Then objTextStream.WriteLine "EXP - 3 - HEAT"
			If caf_exp_pay_ac_checkbox = checked 			Then objTextStream.WriteLine "EXP - 3 - ACON"
			If caf_exp_pay_electricity_checkbox = checked 	Then objTextStream.WriteLine "EXP - 3 - ELEC"
			If caf_exp_pay_phone_checkbox = checked 		Then objTextStream.WriteLine "EXP - 3 - PHON"
			If caf_exp_pay_none_checkbox = checked 			Then objTextStream.WriteLine "EXP - 3 - NONE"
			objTextStream.WriteLine "EXP - 4 - " & exp_migrant_seasonal_formworker_yn
			objTextStream.WriteLine "EXP - 5 - PREV - " & exp_received_previous_assistance_yn
			objTextStream.WriteLine "EXP - 5 - WHEN - " & exp_previous_assistance_when
			objTextStream.WriteLine "EXP - 5 - WHER - " & exp_previous_assistance_where
			objTextStream.WriteLine "EXP - 5 - WHAT - " & exp_previous_assistance_what
			objTextStream.WriteLine "EXP - 6 - PREG - " & exp_pregnant_yn
			objTextStream.WriteLine "EXP - 6 - WHO? - " & exp_pregnant_who
			objTextStream.WriteLine "EXP - INTVW - INCM - " & intv_app_month_income
			objTextStream.WriteLine "EXP - INTVW - ASST - " & intv_app_month_asset
			objTextStream.WriteLine "EXP - INTVW - RENT - " & intv_app_month_housing_expense
			If intv_exp_pay_heat_checkbox = checked 		Then objTextStream.WriteLine "EXP - INTVW - HEAT"
			If intv_exp_pay_ac_checkbox = checked 			Then objTextStream.WriteLine "EXP - INTVW - ACON"
			If intv_exp_pay_electricity_checkbox = checked 	Then objTextStream.WriteLine "EXP - INTVW - ELEC"
			If intv_exp_pay_phone_checkbox = checked 		Then objTextStream.WriteLine "EXP - INTVW - PHON"
			If intv_exp_pay_none_checkbox = checked 		Then objTextStream.WriteLine "EXP - INTVW - NONE"
			objTextStream.WriteLine "EXP - INTVW - ID - " & id_verif_on_file
			objTextStream.WriteLine "EXP - INTVW - 89 - " & snap_active_in_other_state
			objTextStream.WriteLine "EXP - INTVW - EXP - " & last_snap_was_exp

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
			objTextStream.WriteLine "01I - " & question_1_interview_notes

			objTextStream.WriteLine "02A - " & question_2_yn
			objTextStream.WriteLine "02N - " & question_2_notes
			objTextStream.WriteLine "02V - " & question_2_verif_yn
			objTextStream.WriteLine "02D - " & question_2_verif_details
			objTextStream.WriteLine "02I - " & question_2_interview_notes

			objTextStream.WriteLine "03A - " & question_3_yn
			objTextStream.WriteLine "03N - " & question_3_notes
			objTextStream.WriteLine "03V - " & question_3_verif_yn
			objTextStream.WriteLine "03D - " & question_3_verif_details
			objTextStream.WriteLine "03I - " & question_3_interview_notes

			objTextStream.WriteLine "04A - " & question_4_yn
			objTextStream.WriteLine "04N - " & question_4_notes
			objTextStream.WriteLine "04V - " & question_4_verif_yn
			objTextStream.WriteLine "04D - " & question_4_verif_details
			objTextStream.WriteLine "04I - " & question_4_interview_notes

			objTextStream.WriteLine "05A - " & question_5_yn
			objTextStream.WriteLine "05N - " & question_5_notes
			objTextStream.WriteLine "05V - " & question_5_verif_yn
			objTextStream.WriteLine "05D - " & question_5_verif_details
			objTextStream.WriteLine "05I - " & question_5_interview_notes

			objTextStream.WriteLine "06A - " & question_6_yn
			objTextStream.WriteLine "06N - " & question_6_notes
			objTextStream.WriteLine "06V - " & question_6_verif_yn
			objTextStream.WriteLine "06D - " & question_6_verif_details
			objTextStream.WriteLine "06I - " & question_6_interview_notes

			objTextStream.WriteLine "07A - " & question_7_yn
			objTextStream.WriteLine "07N - " & question_7_notes
			objTextStream.WriteLine "07V - " & question_7_verif_yn
			objTextStream.WriteLine "07D - " & question_7_verif_details
			objTextStream.WriteLine "07I - " & question_7_interview_notes

			objTextStream.WriteLine "08A - " & question_8_yn
			objTextStream.WriteLine "08N - " & question_8_notes
			objTextStream.WriteLine "08V - " & question_8_verif_yn
			objTextStream.WriteLine "08D - " & question_8_verif_details
			objTextStream.WriteLine "08I - " & question_8_interview_notes

			objTextStream.WriteLine "09A - " & question_9_yn
			objTextStream.WriteLine "09N - " & question_9_notes
			objTextStream.WriteLine "09V - " & question_9_verif_yn
			objTextStream.WriteLine "09D - " & question_9_verif_details

			objTextStream.WriteLine "10A - " & question_10_yn
			objTextStream.WriteLine "10N - " & question_10_notes
			objTextStream.WriteLine "10V - " & question_10_verif_yn
			objTextStream.WriteLine "10D - " & question_10_verif_details
			objTextStream.WriteLine "10G - " & question_10_monthly_earnings
			objTextStream.WriteLine "10I - " & question_10_interview_notes

			objTextStream.WriteLine "11A - " & question_11_yn
			objTextStream.WriteLine "11N - " & question_11_notes
			objTextStream.WriteLine "11V - " & question_11_verif_yn
			objTextStream.WriteLine "11D - " & question_11_verif_details
			objTextStream.WriteLine "11I - " & question_11_interview_notes

			objTextStream.WriteLine "PWE - " & pwe_selection

			objTextStream.WriteLine "12A - RS - " & question_12_rsdi_yn
			objTextStream.WriteLine "12$ - RS - " & question_12_rsdi_amt
			objTextStream.WriteLine "12A - SS - " & question_12_ssi_yn
			objTextStream.WriteLine "12$ - SS - " & question_12_ssi_amt
			objTextStream.WriteLine "12A - VA - " & question_12_va_yn
			objTextStream.WriteLine "12$ - VA - " & question_12_va_amt
			objTextStream.WriteLine "12A - UI - " & question_12_ui_yn
			objTextStream.WriteLine "12$ - UI - " & question_12_ui_amt
			objTextStream.WriteLine "12A - WC - " & question_12_wc_yn
			objTextStream.WriteLine "12$ - WC - " & question_12_wc_amt
			objTextStream.WriteLine "12A - RT - " & question_12_ret_yn
			objTextStream.WriteLine "12$ - RT - " & question_12_ret_amt
			objTextStream.WriteLine "12A - TP - " & question_12_trib_yn
			objTextStream.WriteLine "12$ - TP - " & question_12_trib_amt
			objTextStream.WriteLine "12A - CS - " & question_12_cs_yn
			objTextStream.WriteLine "12$ - CS - " & question_12_cs_amt
			objTextStream.WriteLine "12A - OT - " & question_12_other_yn
			objTextStream.WriteLine "12$ - OT - " & question_12_other_amt
			objTextStream.WriteLine "12A - " & q_12_answered
			objTextStream.WriteLine "12N - " & question_12_notes
			objTextStream.WriteLine "12V - " & question_12_verif_yn
			objTextStream.WriteLine "12D - " & question_12_verif_details
			objTextStream.WriteLine "12I - " & question_12_interview_notes

			objTextStream.WriteLine "13A - " & question_13_yn
			objTextStream.WriteLine "13N - " & question_13_notes
			objTextStream.WriteLine "13V - " & question_13_verif_yn
			objTextStream.WriteLine "13D - " & question_13_verif_details
			objTextStream.WriteLine "13I - " & question_13_interview_notes

			objTextStream.WriteLine "14A - RT - " &  question_14_rent_yn
			objTextStream.WriteLine "14A - SB - " &  question_14_subsidy_yn
			objTextStream.WriteLine "14A - MT - " &  question_14_mortgage_yn
			objTextStream.WriteLine "14A - AS - " &  question_14_association_yn
			objTextStream.WriteLine "14A - IN - " &  question_14_insurance_yn
			objTextStream.WriteLine "14A - RM - " &  question_14_room_yn
			objTextStream.WriteLine "14A - TX - " &  question_14_taxes_yn
			objTextStream.WriteLine "14A - " & q_14_answered
			objTextStream.WriteLine "14N - " & question_14_notes
			objTextStream.WriteLine "14V - " & question_14_verif_yn
			objTextStream.WriteLine "14D - " & question_14_verif_details
			objTextStream.WriteLine "14I - " & question_14_interview_notes

			objTextStream.WriteLine "15A - HA - " & question_15_heat_ac_yn
			objTextStream.WriteLine "15A - EL - " & question_15_electricity_yn
			objTextStream.WriteLine "15A - CF - " & question_15_cooking_fuel_yn
			objTextStream.WriteLine "15A - WS - " & question_15_water_and_sewer_yn
			objTextStream.WriteLine "15A - GR - " & question_15_garbage_yn
			objTextStream.WriteLine "15A - PN - " & question_15_phone_yn
			objTextStream.WriteLine "15A - LP - " & question_15_liheap_yn
			objTextStream.WriteLine "15A - " & q_15_answered
			objTextStream.WriteLine "15N - " & question_15_notes
			objTextStream.WriteLine "15V - " & question_15_verif_yn
			objTextStream.WriteLine "15D - " & question_15_verif_details
			objTextStream.WriteLine "15I - " & question_15_interview_notes

			objTextStream.WriteLine "16A - " & question_16_yn
			objTextStream.WriteLine "16N - " & question_16_notes
			objTextStream.WriteLine "16V - " & question_16_verif_yn
			objTextStream.WriteLine "16D - " & question_16_verif_details
			objTextStream.WriteLine "16I - " & question_16_interview_notes

			objTextStream.WriteLine "17A - " & question_17_yn
			objTextStream.WriteLine "17N - " & question_17_notes
			objTextStream.WriteLine "17V - " & question_17_verif_yn
			objTextStream.WriteLine "17D - " & question_17_verif_details
			objTextStream.WriteLine "17I - " & question_17_interview_notes

			objTextStream.WriteLine "18A - " & question_18_yn
			objTextStream.WriteLine "18N - " & question_18_notes
			objTextStream.WriteLine "18V - " & question_18_verif_yn
			objTextStream.WriteLine "18D - " & question_18_verif_details
			objTextStream.WriteLine "18I - " & question_18_interview_notes

			objTextStream.WriteLine "19A - " & question_19_yn
			objTextStream.WriteLine "19N - " & question_19_notes
			objTextStream.WriteLine "19V - " & question_19_verif_yn
			objTextStream.WriteLine "19D - " & question_19_verif_details
			objTextStream.WriteLine "19I - " & question_19_interview_notes

			objTextStream.WriteLine "20A - CA - " & question_20_cash_yn
			objTextStream.WriteLine "20A - AC - " & question_20_acct_yn
			objTextStream.WriteLine "20A - SE - " & question_20_secu_yn
			objTextStream.WriteLine "20A - CR - " & question_20_cars_yn
			objTextStream.WriteLine "20A - " & q_20_answered
			objTextStream.WriteLine "20N - " & question_20_notes
			objTextStream.WriteLine "20V - " & question_20_verif_yn
			objTextStream.WriteLine "20D - " & question_20_verif_details
			objTextStream.WriteLine "20I - " & question_20_interview_notes

			objTextStream.WriteLine "21A - " & question_21_yn
			objTextStream.WriteLine "21N - " & question_21_notes
			objTextStream.WriteLine "21V - " & question_21_verif_yn
			objTextStream.WriteLine "21D - " & question_21_verif_details
			objTextStream.WriteLine "21I - " & question_21_interview_notes

			objTextStream.WriteLine "22A - " & question_22_yn
			objTextStream.WriteLine "22N - " & question_22_notes
			objTextStream.WriteLine "22V - " & question_22_verif_yn
			objTextStream.WriteLine "22D - " & question_22_verif_details
			objTextStream.WriteLine "22I - " & question_22_interview_notes

			objTextStream.WriteLine "23A - " & question_23_yn
			objTextStream.WriteLine "23N - " & question_23_notes
			objTextStream.WriteLine "23V - " & question_23_verif_yn
			objTextStream.WriteLine "23D - " & question_23_verif_details
			objTextStream.WriteLine "23I - " & question_23_interview_notes

			objTextStream.WriteLine "24A - RP - " & question_24_rep_payee_yn
			objTextStream.WriteLine "24A - GF - " & question_24_guardian_fees_yn
			objTextStream.WriteLine "24A - SD - " & question_24_special_diet_yn
			objTextStream.WriteLine "24A - HH - " & question_24_high_housing_yn
			objTextStream.WriteLine "24A - " & q_24_answered
			objTextStream.WriteLine "24N - " & question_24_notes
			objTextStream.WriteLine "24V - " & question_24_verif_yn
			objTextStream.WriteLine "24D - " & question_24_verif_details
			objTextStream.WriteLine "24I - " & question_24_interview_notes

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

			objTextStream.WriteLine "AREP - 01 - " & arep_name
			objTextStream.WriteLine "AREP - 02 - " & arep_relationship
			objTextStream.WriteLine "AREP - 03 - " & arep_phone_number
			objTextStream.WriteLine "AREP - 04 - " & arep_addr_street
			objTextStream.WriteLine "AREP - 05 - " & arep_addr_city
			objTextStream.WriteLine "AREP - 06 - " & arep_addr_state
			objTextStream.WriteLine "AREP - 07 - " & arep_addr_zip
			objTextStream.WriteLine "AREP - 08 - " & arep_complete_forms_checkbox
			objTextStream.WriteLine "AREP - 09 - " & arep_get_notices_checkbox
			objTextStream.WriteLine "AREP - 10 - " & arep_use_SNAP_checkbox
			objTextStream.WriteLine "SIG - 01 - " & signature_person
			objTextStream.WriteLine "SIG - 02 - " & signature_date
			objTextStream.WriteLine "SIG - 03 - " & second_signature_person
			objTextStream.WriteLine "SIG - 04 - " & second_signature_date
			objTextStream.WriteLine "SIG - 05 - " & client_signed_verbally_yn
			objTextStream.WriteLine "SIG - 06 - " & interview_date

			objTextStream.WriteLine "FORM - 01 - " & confirm_resp_read
			objTextStream.WriteLine "FORM - 02 - " & confirm_rights_read
			objTextStream.WriteLine "FORM - 03 - " & confirm_ebt_read
			objTextStream.WriteLine "FORM - 04 - " & confirm_ebt_how_to_read
			objTextStream.WriteLine "FORM - 05 - " & confirm_npp_info_read
			objTextStream.WriteLine "FORM - 06 - " & confirm_npp_rights_read
			objTextStream.WriteLine "FORM - 07 - " & confirm_appeal_rights_read
			objTextStream.WriteLine "FORM - 08 - " & confirm_civil_rights_read
			objTextStream.WriteLine "FORM - 09 - " & confirm_cover_letter_read
			objTextStream.WriteLine "FORM - 10 - " & confirm_program_information_read
			objTextStream.WriteLine "FORM - 11 - " & confirm_DV_read
			objTextStream.WriteLine "FORM - 12 - " & confirm_disa_read
			objTextStream.WriteLine "FORM - 13 - " & confirm_mfip_forms_read
			objTextStream.WriteLine "FORM - 14 - " & confirm_mfip_cs_read
			objTextStream.WriteLine "FORM - 15 - " & confirm_minor_mfip_read
			objTextStream.WriteLine "FORM - 16 - " & confirm_snap_forms_read
			objTextStream.WriteLine "FORM - 17 - " & confirm_recap_read

			For known_membs = 0 to UBound(HH_MEMB_ARRAY, 2)
				' objTextStream.WriteLine "ARR - ALL_CLIENTS_ARRAY - " & ALL_CLIENTS_ARRAY(memb_last_name, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_first_name, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_mid_name, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_other_names, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_ssn_verif, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_soc_sec_numb, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_dob, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_gender, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_rel_to_applct, known_membs)&"~"&ALL_CLIENTS_ARRAY(memi_marriage_status, known_membs)&"~"&ALL_CLIENTS_ARRAY(memi_last_grade, known_membs)&"~"&ALL_CLIENTS_ARRAY(memi_MN_entry_date, known_membs)&"~"&ALL_CLIENTS_ARRAY(memi_former_state, known_membs)&"~"&ALL_CLIENTS_ARRAY(memi_citizen, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_interpreter, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_spoken_language, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_written_language, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_ethnicity, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_race_a_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_race_b_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_race_n_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_race_p_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_race_w_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_snap_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_cash_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_emer_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_none_checkbox, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_intend_to_reside_mn, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_imig_status, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_sponsor_yn, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_verif_yn, known_membs)&"~"&ALL_CLIENTS_ARRAY(clt_verif_details, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_notes, known_membs)&"~"&ALL_CLIENTS_ARRAY(memb_ref_numb, known_membs)
				race_a_info = ""
				race_b_info = ""
				race_n_info = ""
				race_p_info = ""
				race_w_info = ""
				prog_s_info = ""
				prog_c_info = ""
				prog_e_info = ""
				prog_n_info = ""

				If HH_MEMB_ARRAY(race_a_checkbox, known_membs) = checked Then race_a_info = "YES"
				If HH_MEMB_ARRAY(race_b_checkbox, known_membs) = checked Then race_b_info = "YES"
				If HH_MEMB_ARRAY(race_n_checkbox, known_membs) = checked Then race_n_info = "YES"
				If HH_MEMB_ARRAY(race_p_checkbox, known_membs) = checked Then race_p_info = "YES"
				If HH_MEMB_ARRAY(race_w_checkbox, known_membs) = checked Then race_w_info = "YES"
				If HH_MEMB_ARRAY(snap_req_checkbox, known_membs) = checked Then prog_s_info = "YES"
				If HH_MEMB_ARRAY(cash_req_checkbox, known_membs) = checked Then prog_c_info = "YES"
				If HH_MEMB_ARRAY(emer_req_checkbox, known_membs) = checked Then prog_e_info = "YES"
				If HH_MEMB_ARRAY(none_req_checkbox, known_membs) = checked Then prog_n_info = "YES"

				objTextStream.WriteLine "ARR - HH_MEMB_ARRAY - " & HH_MEMB_ARRAY(ref_number, known_membs)&"~"&HH_MEMB_ARRAY(access_denied, known_membs)&"~"&HH_MEMB_ARRAY(full_name_const, known_membs)&"~"&HH_MEMB_ARRAY(last_name_const, known_membs)&"~"&_
				HH_MEMB_ARRAY(first_name_const, known_membs)&"~"&HH_MEMB_ARRAY(mid_initial, known_membs)&"~"&HH_MEMB_ARRAY(age, known_membs)&"~"&HH_MEMB_ARRAY(date_of_birth, known_membs)&"~"&HH_MEMB_ARRAY(ssn, known_membs)&"~"&HH_MEMB_ARRAY(ssn_verif, known_membs)&"~"&_
				HH_MEMB_ARRAY(birthdate_verif, known_membs)&"~"&HH_MEMB_ARRAY(gender, known_membs)&"~"&HH_MEMB_ARRAY(race, known_membs)&"~"&HH_MEMB_ARRAY(spoken_lang, known_membs)&"~"&HH_MEMB_ARRAY(written_lang, known_membs)&"~"&HH_MEMB_ARRAY(interpreter, known_membs)&"~"&_
				HH_MEMB_ARRAY(alias_yn, known_membs)&"~"&HH_MEMB_ARRAY(ethnicity_yn, known_membs)&"~"&HH_MEMB_ARRAY(id_verif, known_membs)&"~"&HH_MEMB_ARRAY(rel_to_applcnt, known_membs)&"~"&HH_MEMB_ARRAY(cash_minor, known_membs)&"~"&HH_MEMB_ARRAY(snap_minor, known_membs)&"~"&_
				HH_MEMB_ARRAY(marital_status, known_membs)&"~"&HH_MEMB_ARRAY(spouse_ref, known_membs)&"~"&HH_MEMB_ARRAY(spouse_name, known_membs)&"~"&HH_MEMB_ARRAY(last_grade_completed, known_membs)&"~"&HH_MEMB_ARRAY(citizen, known_membs)&"~"&_
				HH_MEMB_ARRAY(other_st_FS_end_date, known_membs)&"~"&HH_MEMB_ARRAY(in_mn_12_mo, known_membs)&"~"&HH_MEMB_ARRAY(residence_verif, known_membs)&"~"&HH_MEMB_ARRAY(mn_entry_date, known_membs)&"~"&HH_MEMB_ARRAY(former_state, known_membs)&"~"&_
				HH_MEMB_ARRAY(fs_pwe, known_membs)&"~"&HH_MEMB_ARRAY(button_one, known_membs)&"~"&HH_MEMB_ARRAY(button_two, known_membs)&"~"&HH_MEMB_ARRAY(clt_has_sponsor, known_membs)&"~"&HH_MEMB_ARRAY(client_verification, known_membs)&"~"&_
				HH_MEMB_ARRAY(client_verification_details, known_membs)&"~"&HH_MEMB_ARRAY(client_notes, known_membs)&"~"&HH_MEMB_ARRAY(intend_to_reside_in_mn, known_membs)&"~"&race_a_info&"~"&race_b_info&"~"&race_n_info&"~"&race_p_info&"~"&race_w_info&"~"&prog_s_info&"~"&prog_c_info&"~"&_
				prog_e_info&"~"&prog_n_info&"~"&HH_MEMB_ARRAY(ssn_no_space, known_membs)&"~"&HH_MEMB_ARRAY(edrs_msg, known_membs)&"~"&HH_MEMB_ARRAY(edrs_match, known_membs)&"~"&_
				HH_MEMB_ARRAY(edrs_notes, known_membs)&"~"&HH_MEMB_ARRAY(last_const, known_membs)
			Next

			for this_jobs = 0 to UBOUND(JOBS_ARRAY, 2)
				objTextStream.WriteLine "ARR - JOBS_ARRAY - " & JOBS_ARRAY(jobs_employee_name, this_jobs)&"~"&JOBS_ARRAY(jobs_hourly_wage, this_jobs)&"~"&JOBS_ARRAY(jobs_gross_monthly_earnings, this_jobs)&"~"&_
				JOBS_ARRAY(jobs_employer_name, this_jobs)&"~"&JOBS_ARRAY(jobs_edit_btn, this_jobs)&"~"&JOBS_ARRAY(jobs_intv_notes, this_jobs)&"~"&JOBS_ARRAY(verif_yn, this_jobs)&"~"&JOBS_ARRAY(verif_details, this_jobs)&"~"&JOBS_ARRAY(jobs_notes, this_jobs)
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
	local_changelog_path = user_myDocs_folder & "interview-answers-" & MAXIS_case_number & "-info.txt"

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
					' MsgBox "~" & left(text_line, 9) & "~" & vbCr & text_line
					' MsgBox text_line
					If left(text_line, 9) = "PRE - WHO" Then who_are_we_completing_the_interview_with = Mid(text_line, 13)
					If left(text_line, 9) = "PRE - ATC" Then all_the_clients = Mid(text_line, 13)
					If left(text_line, 7) = "EXP - 1" Then exp_q_1_income_this_month = Mid(text_line, 11)
					If left(text_line, 7) = "EXP - 2" Then exp_q_2_assets_this_month = Mid(text_line, 11)
					If left(text_line, 14) = "EXP - 3 - RENT" Then exp_q_3_rent_this_month = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 3 - HEAT" Then caf_exp_pay_heat_checkbox = checked
					If left(text_line, 14) = "EXP - 3 - ACON" Then caf_exp_pay_ac_checkbox = checked
					If left(text_line, 14) = "EXP - 3 - ELEC" Then caf_exp_pay_electricity_checkbox = checked
					If left(text_line, 14) = "EXP - 3 - PHON" Then caf_exp_pay_phone_checkbox = checked
					If left(text_line, 14) = "EXP - 3 - NONE" Then caf_exp_pay_none_checkbox = checked
					If left(text_line, 7) = "EXP - 4" Then exp_migrant_seasonal_formworker_yn = Mid(text_line, 11)
					If left(text_line, 14) = "EXP - 5 - PREV" Then exp_received_previous_assistance_yn = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 5 - WHEN" Then exp_previous_assistance_when = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 5 - WHER" Then exp_previous_assistance_where = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 5 - WHAT" Then exp_previous_assistance_what = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 6 - PREG" Then exp_pregnant_yn = Mid(text_line, 18)
					If left(text_line, 14) = "EXP - 6 - WHO?" Then exp_pregnant_who = Mid(text_line, 18)

					If left(text_line, 18) = "EXP - INTVW - INCM" Then intv_app_month_income = Mid(text_line, 22)
					If left(text_line, 18) = "EXP - INTVW - ASST" Then intv_app_month_asset = Mid(text_line, 22)
					If left(text_line, 18) = "EXP - INTVW - RENT" Then intv_app_month_housing_expense = Mid(text_line, 22)
					If left(text_line, 18) = "EXP - INTVW - HEAT" Then intv_exp_pay_heat_checkbox = checked
					If left(text_line, 18) = "EXP - INTVW - ACON" Then intv_exp_pay_ac_checkbox = checked
					If left(text_line, 18) = "EXP - INTVW - ELEC" Then intv_exp_pay_electricity_checkbox = checked
					If left(text_line, 18) = "EXP - INTVW - PHON" Then intv_exp_pay_phone_checkbox = checked
					If left(text_line, 18) = "EXP - INTVW - NONE" Then intv_exp_pay_none_checkbox = checked
					If left(text_line, 16) = "EXP - INTVW - ID" Then id_verif_on_file = Mid(text_line, 20)
					If left(text_line, 16) = "EXP - INTVW - 89" Then snap_active_in_other_state = Mid(text_line, 20)
					If left(text_line, 17) = "EXP - INTVW - EXP" Then last_snap_was_exp = Mid(text_line, 21)

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
					If left(text_line, 3) = "01I" Then question_1_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "02A" Then question_2_yn = Mid(text_line, 7)
					If left(text_line, 3) = "02N" Then question_2_notes = Mid(text_line, 7)
					If left(text_line, 3) = "02V" Then question_2_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "02D" Then question_2_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "02I" Then question_2_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "03A" Then question_3_yn = Mid(text_line, 7)
					If left(text_line, 3) = "03N" Then question_3_notes = Mid(text_line, 7)
					If left(text_line, 3) = "03V" Then question_3_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "03D" Then question_3_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "03I" Then question_3_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "04A" Then question_4_yn = Mid(text_line, 7)
					If left(text_line, 3) = "04N" Then question_4_notes = Mid(text_line, 7)
					If left(text_line, 3) = "04V" Then question_4_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "04D" Then question_4_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "04I" Then question_4_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "05A" Then question_5_yn = Mid(text_line, 7)
					If left(text_line, 3) = "05N" Then question_5_notes = Mid(text_line, 7)
					If left(text_line, 3) = "05V" Then question_5_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "05D" Then question_5_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "05I" Then question_5_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "06A" Then question_6_yn = Mid(text_line, 7)
					If left(text_line, 3) = "06N" Then question_6_notes = Mid(text_line, 7)
					If left(text_line, 3) = "06V" Then question_6_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "06D" Then question_6_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "06I" Then question_6_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "07A" Then question_7_yn = Mid(text_line, 7)
					If left(text_line, 3) = "07N" Then question_7_notes = Mid(text_line, 7)
					If left(text_line, 3) = "07V" Then question_7_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "07D" Then question_7_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "07I" Then question_7_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "08A" Then question_8_yn = Mid(text_line, 7)
					If left(text_line, 3) = "08N" Then question_8_notes = Mid(text_line, 7)
					If left(text_line, 3) = "08V" Then question_8_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "08D" Then question_8_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "08I" Then question_8_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "09A" Then question_9_yn = Mid(text_line, 7)
					If left(text_line, 3) = "09N" Then question_9_notes = Mid(text_line, 7)
					If left(text_line, 3) = "09V" Then question_9_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "09D" Then question_9_verif_details = Mid(text_line, 7)

					If left(text_line, 3) = "10A" Then question_10_yn = Mid(text_line, 7)
					If left(text_line, 3) = "10N" Then question_10_notes = Mid(text_line, 7)
					If left(text_line, 3) = "10V" Then question_10_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "10D" Then question_10_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "10G" Then question_10_monthly_earnings = Mid(text_line, 7)
					If left(text_line, 3) = "10I" Then question_10_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "11A" Then question_11_yn = Mid(text_line, 7)
					If left(text_line, 3) = "11N" Then question_11_notes = Mid(text_line, 7)
					If left(text_line, 3) = "11V" Then question_11_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "11D" Then question_11_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "11I" Then question_11_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "PWE" Then pwe_selection = Mid(text_line, 7)

					If left(text_line, 8) = "12A - RS" Then question_12_rsdi_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12$ - RS" Then question_12_rsdi_amt = Mid(text_line, 12)
					If left(text_line, 8) = "12A - SS" Then question_12_ssi_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12$ - SS" Then question_12_ssi_amt = Mid(text_line, 12)
					If left(text_line, 8) = "12A - VA" Then question_12_va_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12$ - VA" Then question_12_va_amt = Mid(text_line, 12)
					If left(text_line, 8) = "12A - UI" Then question_12_ui_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12$ - UI" Then question_12_ui_amt = Mid(text_line, 12)
					If left(text_line, 8) = "12A - WC" Then question_12_wc_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12$ - WC" Then question_12_wc_amt = Mid(text_line, 12)
					If left(text_line, 8) = "12A - RT" Then question_12_ret_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12$ - RT" Then question_12_ret_amt = Mid(text_line, 12)
					If left(text_line, 8) = "12A - TP" Then question_12_trib_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12$ - TP" Then question_12_trib_amt = Mid(text_line, 12)
					If left(text_line, 8) = "12A - CS" Then question_12_cs_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12$ - CS" Then question_12_cs_amt = Mid(text_line, 12)
					If left(text_line, 8) = "12A - OT" Then question_12_other_yn = Mid(text_line, 12)
					If left(text_line, 8) = "12$ - OT" Then question_12_other_amt = Mid(text_line, 12)
					If left(text_line, 3) = "12A" Then q_12_answered = Mid(text_line, 7)
					If left(text_line, 3) = "12N" Then question_12_notes = Mid(text_line, 7)
					If left(text_line, 3) = "12V" Then question_12_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "12D" Then question_12_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "12I" Then question_12_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "13A" Then question_13_yn = Mid(text_line, 7)
					If left(text_line, 3) = "13N" Then question_13_notes = Mid(text_line, 7)
					If left(text_line, 3) = "13V" Then question_13_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "13D" Then question_13_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "13I" Then question_13_interview_notes = Mid(text_line, 7)

					If left(text_line, 8) = "14A - RT" Then  question_14_rent_yn = Mid(text_line, 12)
					If left(text_line, 8) = "14A - SB" Then  question_14_subsidy_yn = Mid(text_line, 12)
					If left(text_line, 8) = "14A - MT" Then  question_14_mortgage_yn = Mid(text_line, 12)
					If left(text_line, 8) = "14A - AS" Then  question_14_association_yn = Mid(text_line, 12)
					If left(text_line, 8) = "14A - IN" Then  question_14_insurance_yn = Mid(text_line, 12)
					If left(text_line, 8) = "14A - RM" Then  question_14_room_yn = Mid(text_line, 12)
					If left(text_line, 8) = "14A - TX" Then  question_14_taxes_yn = Mid(text_line, 12)
					If left(text_line, 3) = "14A" Then q_14_answered = Mid(text_line, 7)
					If left(text_line, 3) = "14N" Then question_14_notes = Mid(text_line, 7)
					If left(text_line, 3) = "14V" Then question_14_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "14D" Then question_14_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "14I" Then question_14_interview_notes = Mid(text_line, 7)

					If left(text_line, 8) = "15A - HA" Then question_15_heat_ac_yn = Mid(text_line, 12)
					If left(text_line, 8) = "15A - EL" Then question_15_electricity_yn = Mid(text_line, 12)
					If left(text_line, 8) = "15A - CF" Then question_15_cooking_fuel_yn = Mid(text_line, 12)
					If left(text_line, 8) = "15A - WS" Then question_15_water_and_sewer_yn = Mid(text_line, 12)
					If left(text_line, 8) = "15A - GR" Then question_15_garbage_yn = Mid(text_line, 12)
					If left(text_line, 8) = "15A - PN" Then question_15_phone_yn = Mid(text_line, 12)
					If left(text_line, 8) = "15A - LP" Then question_15_liheap_yn = Mid(text_line, 12)
					If left(text_line, 3) = "15A" Then q_15_answered = Mid(text_line, 7)
					If left(text_line, 3) = "15N" Then question_15_notes = Mid(text_line, 7)
					If left(text_line, 3) = "15V" Then question_15_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "15D" Then question_15_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "15I" Then question_15_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "16A" Then question_16_yn = Mid(text_line, 7)
					If left(text_line, 3) = "16N" Then question_16_notes = Mid(text_line, 7)
					If left(text_line, 3) = "16V" Then question_16_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "16D" Then question_16_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "16I" Then question_16_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "17A" Then question_17_yn = Mid(text_line, 7)
					If left(text_line, 3) = "17N" Then question_17_notes = Mid(text_line, 7)
					If left(text_line, 3) = "17V" Then question_17_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "17D" Then question_17_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "17I" Then question_17_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "18A" Then question_18_yn = Mid(text_line, 7)
					If left(text_line, 3) = "18N" Then question_18_notes = Mid(text_line, 7)
					If left(text_line, 3) = "18V" Then question_18_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "18D" Then question_18_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "18I" Then question_18_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "19A" Then question_19_yn = Mid(text_line, 7)
					If left(text_line, 3) = "19N" Then question_19_notes = Mid(text_line, 7)
					If left(text_line, 3) = "19V" Then question_19_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "19D" Then question_19_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "19I" Then question_19_interview_notes = Mid(text_line, 7)

					If left(text_line, 8) = "20A - CA" Then question_20_cash_yn = Mid(text_line, 12)
					If left(text_line, 8) = "20A - AC" Then question_20_acct_yn = Mid(text_line, 12)
					If left(text_line, 8) = "20A - SE" Then question_20_secu_yn = Mid(text_line, 12)
					If left(text_line, 8) = "20A - CR" Then question_20_cars_yn = Mid(text_line, 12)
					If left(text_line, 3) = "20A" Then q_20_answered = Mid(text_line, 7)
					If left(text_line, 3) = "20N" Then question_20_notes = Mid(text_line, 7)
					If left(text_line, 3) = "20V" Then question_20_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "20D" Then question_20_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "20I" Then question_20_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "21A" Then question_21_yn = Mid(text_line, 7)
					If left(text_line, 3) = "21N" Then question_21_notes = Mid(text_line, 7)
					If left(text_line, 3) = "21V" Then question_21_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "21D" Then question_21_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "21I" Then question_21_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "22A" Then question_22_yn = Mid(text_line, 7)
					If left(text_line, 3) = "22N" Then question_22_notes = Mid(text_line, 7)
					If left(text_line, 3) = "22V" Then question_22_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "22D" Then question_22_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "22I" Then question_22_interview_notes = Mid(text_line, 7)

					If left(text_line, 3) = "23A" Then question_23_yn = Mid(text_line, 7)
					If left(text_line, 3) = "23N" Then question_23_notes = Mid(text_line, 7)
					If left(text_line, 3) = "23V" Then question_23_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "23D" Then question_23_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "23I" Then question_23_interview_notes = Mid(text_line, 7)

					If left(text_line, 8) = "24A - RP" Then question_24_rep_payee_yn = Mid(text_line, 12)
					If left(text_line, 8) = "24A - GF" Then question_24_guardian_fees_yn = Mid(text_line, 12)
					If left(text_line, 8) = "24A - SD" Then question_24_special_diet_yn = Mid(text_line, 12)
					If left(text_line, 8) = "24A - HH" Then question_24_high_housing_yn = Mid(text_line, 12)
					If left(text_line, 3) = "24A" Then q_24_answered = Mid(text_line, 7)
					If left(text_line, 3) = "24N" Then question_24_notes = Mid(text_line, 7)
					If left(text_line, 3) = "24V" Then question_24_verif_yn = Mid(text_line, 7)
					If left(text_line, 3) = "24D" Then question_24_verif_details = Mid(text_line, 7)
					If left(text_line, 3) = "24I" Then question_24_interview_notes = Mid(text_line, 7)

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

					If left(text_line, 9) = "AREP - 01" Then arep_name = Mid(text_line, 13)
					If left(text_line, 9) = "AREP - 02" Then arep_relationship = Mid(text_line, 13)
					If left(text_line, 9) = "AREP - 03" Then arep_phone_number = Mid(text_line, 13)
					If left(text_line, 9) = "AREP - 04" Then arep_addr_street = Mid(text_line, 13)
					If left(text_line, 9) = "AREP - 05" Then arep_addr_city = Mid(text_line, 13)
					If left(text_line, 9) = "AREP - 06" Then arep_addr_state = Mid(text_line, 13)
					If left(text_line, 9) = "AREP - 07" Then arep_addr_zip = Mid(text_line, 13)
					If left(text_line, 9) = "AREP - 08" Then arep_complete_forms_checkbox = Mid(text_line, 13)
					If left(text_line, 9) = "AREP - 09" Then arep_get_notices_checkbox = Mid(text_line, 13)
					If left(text_line, 9) = "AREP - 10" Then arep_use_SNAP_checkbox = Mid(text_line, 13)

					If left(text_line, 8) = "SIG - 01" Then signature_person = Mid(text_line, 12)
					If left(text_line, 8) = "SIG - 02" Then signature_date = Mid(text_line, 12)
					If left(text_line, 8) = "SIG - 03" Then second_signature_person = Mid(text_line, 12)
					If left(text_line, 8) = "SIG - 04" Then second_signature_date = Mid(text_line, 12)
					If left(text_line, 8) = "SIG - 05" Then client_signed_verbally_yn = Mid(text_line, 12)
					If left(text_line, 8) = "SIG - 06" Then interview_date = Mid(text_line, 12)

					If left(text_line, 9) = "FORM - 01" Then confirm_resp_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 02" Then confirm_rights_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 03" Then confirm_ebt_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 04" Then confirm_ebt_how_to_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 05" Then confirm_npp_info_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 06" Then confirm_npp_rights_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 07" Then confirm_appeal_rights_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 08" Then confirm_civil_rights_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 09" Then confirm_cover_letter_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 10" Then confirm_program_information_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 11" Then confirm_DV_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 12" Then confirm_disa_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 13" Then confirm_mfip_forms_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 14" Then confirm_mfip_cs_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 15" Then confirm_minor_mfip_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 16" Then confirm_snap_forms_read = Mid(text_line, 13)
					If left(text_line, 9) = "FORM - 17" Then confirm_recap_read = Mid(text_line, 13)


					' If left(text_line, 4) = "QQ1A" Then qual_question_one = Mid(text_line, 8)

					If left(text_line, 3) = "ARR" Then
						If MID(text_line, 7, 13) = "HH_MEMB_ARRAY" Then
							array_info = Mid(text_line, 23)
							array_info = split(array_info, "~")
							ReDim Preserve HH_MEMB_ARRAY(last_const, known_membs)
							HH_MEMB_ARRAY(ref_number, known_membs)					= array_info(0)
							HH_MEMB_ARRAY(access_denied, known_membs)				= array_info(1)
							HH_MEMB_ARRAY(full_name_const, known_membs)				= array_info(2)
							HH_MEMB_ARRAY(last_name_const, known_membs)				= array_info(3)
							HH_MEMB_ARRAY(first_name_const, known_membs)			= array_info(4)
							HH_MEMB_ARRAY(mid_initial, known_membs)					= array_info(5)
							HH_MEMB_ARRAY(age, known_membs)							= array_info(6)
							HH_MEMB_ARRAY(date_of_birth, known_membs)				= array_info(7)
							HH_MEMB_ARRAY(ssn, known_membs)							= array_info(8)
							HH_MEMB_ARRAY(ssn_verif, known_membs)					= array_info(9)
							HH_MEMB_ARRAY(birthdate_verif, known_membs)				= array_info(10)
							HH_MEMB_ARRAY(gender, known_membs)						= array_info(11)
							HH_MEMB_ARRAY(race, known_membs)						= array_info(12)
							HH_MEMB_ARRAY(spoken_lang, known_membs)					= array_info(13)
							HH_MEMB_ARRAY(written_lang, known_membs)				= array_info(14)
							HH_MEMB_ARRAY(interpreter, known_membs)					= array_info(15)
							HH_MEMB_ARRAY(alias_yn, known_membs)					= array_info(16)
							HH_MEMB_ARRAY(ethnicity_yn, known_membs)				= array_info(17)
							HH_MEMB_ARRAY(id_verif, known_membs)					= array_info(18)
							HH_MEMB_ARRAY(rel_to_applcnt, known_membs)				= array_info(19)
							HH_MEMB_ARRAY(cash_minor, known_membs)					= array_info(20)
							HH_MEMB_ARRAY(snap_minor, known_membs)					= array_info(21)
							HH_MEMB_ARRAY(marital_status, known_membs)				= array_info(22)
							HH_MEMB_ARRAY(spouse_ref, known_membs)					= array_info(23)
							HH_MEMB_ARRAY(spouse_name, known_membs)					= array_info(24)
							HH_MEMB_ARRAY(last_grade_completed, known_membs) 		= array_info(25)
							HH_MEMB_ARRAY(citizen, known_membs)						= array_info(26)
							HH_MEMB_ARRAY(other_st_FS_end_date, known_membs) 		= array_info(27)
							HH_MEMB_ARRAY(in_mn_12_mo, known_membs)					= array_info(28)
							HH_MEMB_ARRAY(residence_verif, known_membs)				= array_info(29)
							HH_MEMB_ARRAY(mn_entry_date, known_membs)				= array_info(30)
							HH_MEMB_ARRAY(former_state, known_membs)				= array_info(31)
							HH_MEMB_ARRAY(fs_pwe, known_membs)						= array_info(32)
							HH_MEMB_ARRAY(button_one, known_membs)					= array_info(33)
							HH_MEMB_ARRAY(button_two, known_membs)					= array_info(34)
							HH_MEMB_ARRAY(clt_has_sponsor, known_membs)				= array_info(35)
							HH_MEMB_ARRAY(client_verification, known_membs)			= array_info(36)
							HH_MEMB_ARRAY(client_verification_details, known_membs)	= array_info(37)
							HH_MEMB_ARRAY(client_notes, known_membs)				= array_info(38)
							HH_MEMB_ARRAY(intend_to_reside_in_mn, known_membs)		= array_info(39)
							If array_info(40) = "YES" Then HH_MEMB_ARRAY(race_a_checkbox, known_membs) = checked
							If array_info(41) = "YES" Then HH_MEMB_ARRAY(race_b_checkbox, known_membs) = checked
							If array_info(42) = "YES" Then HH_MEMB_ARRAY(race_n_checkbox, known_membs) = checked
							If array_info(43) = "YES" Then HH_MEMB_ARRAY(race_p_checkbox, known_membs) = checked
							If array_info(44) = "YES" Then HH_MEMB_ARRAY(race_w_checkbox, known_membs) = checked
							If array_info(45) = "YES" Then HH_MEMB_ARRAY(snap_req_checkbox, known_membs) = checked
							If array_info(46) = "YES" Then HH_MEMB_ARRAY(cash_req_checkbox, known_membs) = checked
							If array_info(47) = "YES" Then HH_MEMB_ARRAY(emer_req_checkbox, known_membs) = checked
							If array_info(48) = "YES" Then HH_MEMB_ARRAY(none_req_checkbox, known_membs) = checked
							HH_MEMB_ARRAY(ssn_no_space, known_membs)				= array_info(49)
							HH_MEMB_ARRAY(edrs_msg, known_membs)					= array_info(50)
							HH_MEMB_ARRAY(edrs_match, known_membs)					= array_info(51)
							HH_MEMB_ARRAY(edrs_notes, known_membs) 					= array_info(52)
							HH_MEMB_ARRAY(last_const, known_membs)					= array_info(53)

							known_membs = known_membs + 1
						End If

						If MID(text_line, 7, 10) = "JOBS_ARRAY" Then
							array_info = Mid(text_line, 20)
							array_info = split(array_info, "~")
							ReDim Preserve JOBS_ARRAY(jobs_notes, known_jobs)
							JOBS_ARRAY(jobs_employee_name, known_jobs) 			= array_info(0)
							JOBS_ARRAY(jobs_hourly_wage, known_jobs) 			= array_info(1)
							JOBS_ARRAY(jobs_gross_monthly_earnings, known_jobs)	= array_info(2)
							JOBS_ARRAY(jobs_employer_name, known_jobs) 			= array_info(3)
							JOBS_ARRAY(jobs_edit_btn, known_jobs)				= array_info(4)
							JOBS_ARRAY(jobs_intv_notes, known_jobs)				= array_info(5)
							JOBS_ARRAY(verif_yn, known_jobs)					= array_info(6)
							JOBS_ARRAY(verif_details, known_jobs)				= array_info(7)
							JOBS_ARRAY(jobs_notes, known_jobs) 					= array_info(8)
							known_jobs = known_jobs + 1
						End If
					End If
				Next
			End If
		End If
	End With
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

		' HH_MEMB_ARRAY(clt_count).intend_to_reside_in_mn = "Yes"

		' ReDim Preserve ALL_ANSWERS_ARRAY(ans_notes, clt_count)
		clt_count = clt_count + 1
	Next

	For i = 0 to UBOUND(HH_MEMB_ARRAY, 2)
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
	For i = 0 to UBOUND(HH_MEMB_ARRAY, 2)
		If HH_MEMB_ARRAY(i).rel_to_applcnt <> "Self" AND HH_MEMB_ARRAY(i).rel_to_applcnt <> "Not Related" AND HH_MEMB_ARRAY(i).rel_to_applcnt <> "Live-in Attendant" AND HH_MEMB_ARRAY(i).rel_to_applcnt <> "Unknown" Then
			ReDim Preserve ALL_HH_RELATIONSHIPS_ARRAY(rela_notes, rela_counter)

			ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_ref, rela_counter) = HH_MEMB_ARRAY(i).ref_number
			ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_name, rela_counter) = HH_MEMB_ARRAY(i).full_name
			ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_ref, rela_counter) = "01"
			ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_name, rela_counter) = HH_MEMB_ARRAY(full_name_const, 0)
			ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = HH_MEMB_ARRAY(i).rel_to_applcnt

			rela_counter = rela_counter + 1

			ReDim Preserve ALL_HH_RELATIONSHIPS_ARRAY(rela_notes, rela_counter)

		 	' MsgBox "Member Count - " & i & vbNewLine & "Relationship - " & HH_MEMB_ARRAY(i).rel_to_applcnt
			ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_ref, rela_counter) = "01"
			ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_name, rela_counter) = HH_MEMB_ARRAY(full_name_const, 0)
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
				If HH_MEMB_ARRAY(gender, 0) = "Female" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Neice"
				If HH_MEMB_ARRAY(gender, 0) = "Female" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Nephew"
			End If
			If HH_MEMB_ARRAY(i).rel_to_applcnt = "Nephew" OR HH_MEMB_ARRAY(i).rel_to_applcnt = "Neice" Then
				If HH_MEMB_ARRAY(gender, 0) = "Female" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Aunt"
				If HH_MEMB_ARRAY(gender, 0) = "Female" Then ALL_HH_RELATIONSHIPS_ARRAY(rela_type, rela_counter) = "Uncle"
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

						For i = 0 to UBOUND(HH_MEMB_ARRAY, 2)
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



	For i = 0 to UBound(HH_MEMB_ARRAY, 2)						'we start with 1 because 0 is MEMB 01 and that parental relationshipare all known because of MEMB
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

					for x = 0 to UBound(HH_MEMB_ARRAY, 2)
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

					for x = 0 to UBound(HH_MEMB_ARRAY, 2)
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
		Case 25
			verif_selection = JOBS_ARRAY(verif_yn, this_jobs)
			verif_detials = JOBS_ARRAY(verif_details, this_jobs)
			question_words = "9.  Does anyone in the household have a job or expect to get income from a job this month or next month? Enter verification for "	& JOBS_ARRAY(jobs_employer_name, this_jobs)
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
		Case 25
			JOBS_ARRAY(verif_yn, this_jobs) = verif_selection
			JOBS_ARRAY(verif_details, this_jobs) = verif_detials
	End Select

end function

function jobs_details_dlg(this_jobs)
	Do
		pick_a_client = replace(all_the_clients, "Select or Type", "Select One...")
		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 321, 165, "Add Job"
		  DropListBox 10, 35, 135, 45, pick_a_client+chr(9)+"", JOBS_ARRAY(jobs_employee_name, this_jobs)
		  EditBox 150, 35, 60, 15, JOBS_ARRAY(jobs_hourly_wage, this_jobs)
		  EditBox 215, 35, 100, 15, JOBS_ARRAY(jobs_gross_monthly_earnings, this_jobs)
		  EditBox 10, 65, 305, 15, JOBS_ARRAY(jobs_employer_name, this_jobs)
		  EditBox 10, 95, 305, 15, JOBS_ARRAY(jobs_notes, this_jobs)
		  EditBox 10, 125, 305, 15, JOBS_ARRAY(jobs_intv_notes, this_jobs)

		  ButtonGroup ButtonPressed
		    PushButton 265, 145, 50, 15, "Return", return_btn
			PushButton 120, 150, 75, 10, "ADD VERIFICATION", add_verif_jobs_btn
		    PushButton 265, 10, 50, 10, "CLEAR", clear_job_btn
		  Text 10, 10, 100, 10, "Enter Job Details/Information"
		  Text 10, 25, 70, 10, "EMPLOYEE NAME:"
		  Text 150, 25, 60, 10, "HOURLY WAGE:"
		  Text 215, 25, 105, 10, "GROSS MONTHLY EARNINGS:"
		  Text 10, 55, 110, 10, "EMPLOYER/BUSINESS NAME:"
		  Text 10, 85, 110, 10, "CAF WRITE-IN INFORMATION:"
		  Text 10, 115, 85, 10, "INTERVIEW NOTES:"
		  Text 10, 150, 110, 10, "JOB Verification - " & JOBS_ARRAY(verif_yn, this_jobs)
		EndDialog


		dialog Dialog1
		If ButtonPressed = -1 Then ButtonPressed = return_btn
		If ButtonPressed = add_verif_jobs_btn Then Call verif_details_dlg(25)
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
const jobs_intv_notes				= 5
const verif_yn						= 6
const verif_details					= 7
const jobs_notes 					= 8

Const end_of_doc = 6			'This is for word document ennumeration

Call find_user_name(worker_name)						'defaulting the name of the suer running the script
Dim TABLE_ARRAY
Dim ALL_CLIENTS_ARRAY
Dim JOBS_ARRAY
ReDim ALL_CLIENTS_ARRAY(memb_notes, 0)
ReDim JOBS_ARRAY(jobs_notes, 0)

Call remove_dash_from_droplist(state_list)
'These are all the definitions for droplists

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

marital_status_list = "Select One..."
marital_status_list = marital_status_list+chr(9)+"N  Never Married"
marital_status_list = marital_status_list+chr(9)+"M  Married Living With Spouse"
marital_status_list = marital_status_list+chr(9)+"S  Married Living Apart (Sep)"
marital_status_list = marital_status_list+chr(9)+"L  Legally Sep"
marital_status_list = marital_status_list+chr(9)+"D  Divorced"
marital_status_list = marital_status_list+chr(9)+"W  Widowed"

question_answers = ""+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Blank"

'Dimming all the variables because they are defined and set within functions
Dim who_are_we_completing_the_interview_with, caf_person_one, exp_q_1_income_this_month, exp_q_2_assets_this_month, exp_q_3_rent_this_month, caf_exp_pay_heat_checkbox, caf_exp_pay_ac_checkbox, caf_exp_pay_electricity_checkbox, caf_exp_pay_phone_checkbox
Dim exp_pay_none_checkbox, exp_migrant_seasonal_formworker_yn, exp_received_previous_assistance_yn, exp_previous_assistance_when, exp_previous_assistance_where, exp_previous_assistance_what, exp_pregnant_yn, exp_pregnant_who, resi_addr_street_full
Dim resi_addr_city, resi_addr_state, resi_addr_zip, reservation_yn, reservation_name, homeless_yn, living_situation, mail_addr_street_full, mail_addr_city, mail_addr_state, mail_addr_zip, phone_one_number, phone_pne_type, phone_two_number
Dim phone_two_type, phone_three_number, phone_three_type, address_change_date, resi_addr_county, caf_form_date, all_the_clients, err_msg
Dim intv_app_month_income, intv_app_month_asset, intv_app_month_housing_expense, intv_exp_pay_heat_checkbox, intv_exp_pay_ac_checkbox, intv_exp_pay_electricity_checkbox, intv_exp_pay_phone_checkbox, intv_exp_pay_none_checkbox
Dim id_verif_on_file, snap_active_in_other_state, last_snap_was_exp

Dim question_1_yn, question_1_notes, question_1_verif_yn, question_1_verif_details, question_1_interview_notes
Dim question_2_yn, question_2_notes, question_2_verif_yn, question_2_verif_details, question_2_interview_notes
Dim question_3_yn, question_3_notes, question_3_verif_yn, question_3_verif_details, question_3_interview_notes
Dim question_4_yn, question_4_notes, question_4_verif_yn, question_4_verif_details, question_4_interview_notes
Dim question_5_yn, question_5_notes, question_5_verif_yn, question_5_verif_details, question_5_interview_notes
Dim question_6_yn, question_6_notes, question_6_verif_yn, question_6_verif_details, question_6_interview_notes
Dim question_7_yn, question_7_notes, question_7_verif_yn, question_7_verif_details, question_7_interview_notes
Dim question_8_yn, question_8a_yn, question_8_notes, question_8_verif_yn, question_8_verif_details, question_8_interview_notes
Dim question_9_yn, question_9_notes, question_9_verif_yn, question_9_verif_details, question_9_interview_notes
Dim question_10_yn, question_10_notes, question_10_verif_yn, question_10_verif_details, question_10_monthly_earnings, question_10_interview_notes
Dim question_11_yn, question_11_notes, question_11_verif_yn, question_11_verif_details, question_11_interview_notes
Dim pwe_selection
Dim question_12_yn, question_12_notes, question_12_verif_yn, question_12_verif_details, question_12_interview_notes
Dim question_12_rsdi_yn, question_12_rsdi_amt, question_12_ssi_yn, question_12_ssi_amt,  question_12_va_yn, question_12_va_amt, question_12_ui_yn, question_12_ui_amt, question_12_wc_yn, question_12_wc_amt, question_12_ret_yn, question_12_ret_amt, question_12_trib_yn, question_12_trib_amt, question_12_cs_yn, question_12_cs_amt, question_12_other_yn, question_12_other_amt
Dim question_13_yn, question_13_notes, question_13_verif_yn, question_13_verif_details, question_13_interview_notes
Dim question_14_yn, question_14_notes, question_14_verif_yn, question_14_verif_details, question_14_interview_notes
Dim question_14_rent_yn, question_14_subsidy_yn, question_14_mortgage_yn, question_14_association_yn, question_14_insurance_yn, question_14_room_yn, question_14_taxes_yn
Dim question_15_yn, question_15_notes, question_15_verif_yn, question_15_verif_details, question_15_interview_notes
Dim question_15_heat_ac_yn, question_15_electricity_yn, question_15_cooking_fuel_yn, question_15_water_and_sewer_yn, question_15_garbage_yn, question_15_phone_yn, question_15_liheap_yn
Dim question_16_yn, question_16_notes, question_16_verif_yn, question_16_verif_details, question_16_interview_notes
Dim question_17_yn, question_17_notes, question_17_verif_yn, question_17_verif_details, question_17_interview_notes
Dim question_18_yn, question_18_notes, question_18_verif_yn, question_18_verif_details, question_18_interview_notes
Dim question_19_yn, question_19_notes, question_19_verif_yn, question_19_verif_details, question_19_interview_notes
Dim question_20_yn, question_20_notes, question_20_verif_yn, question_20_verif_details, question_20_interview_notes
Dim question_20_cash_yn, question_20_acct_yn, question_20_secu_yn, question_20_cars_yn
Dim question_21_yn, question_21_notes, question_21_verif_yn, question_21_verif_details, question_21_interview_notes
Dim question_22_yn, question_22_notes, question_22_verif_yn, question_22_verif_details, question_22_interview_notes
Dim question_23_yn, question_23_notes, question_23_verif_yn, question_23_verif_details, question_23_interview_notes
Dim question_24_yn, question_24_notes, question_24_verif_yn, question_24_verif_details, question_24_interview_notes
Dim question_24_rep_payee_yn, question_24_guardian_fees_yn, question_24_special_diet_yn, question_24_high_housing_yn
Dim qual_question_one, qual_memb_one, qual_question_two, qual_memb_two, qual_question_three, qual_memb_there, qual_question_four, qual_memb_four, qual_question_five, qual_memb_five

Dim show_pg_one_memb01_and_exp, show_pg_one_address, show_pg_memb_list, show_q_1_6
Dim show_q_7_11, show_q_14_15, show_q_21_24, show_qual, show_pg_last

show_pg_one_memb01_and_exp	= 1
show_pg_one_address			= 2
show_pg_memb_list			= 3
show_q_1_6					= 4
show_q_7_11					= 5
show_q_12_13				= 6
show_q_14_15				= 7
show_q_16_20				= 8
show_q_21_24				= 9
show_qual					= 10
show_pg_last				= 11

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
show_err_msg_during_movement = ""

Call back_to_SELF
EMReadScreen MX_region, 10, 22, 48
MX_region = trim(MX_region)
If MX_region = "INQUIRY DB" Then
	continue_in_inquiry = MsgBox("You have started this script run in INQUIRY." & vbNewLine & vbNewLine & "The script cannot complete a CASE:NOTE when run in inquiry. The functionality is limited when run in inquiry. " & vbNewLine & vbNewLine & "Would you like to continue in INQUIRY?", vbQuestion + vbYesNo, "Continue in INQUIRY")
	If continue_in_inquiry = vbNo Then Call script_end_procedure("~PT Interview Script cancelled as it was run in inquiry.")
End If

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 371, 330, "Interview Script Case number dialog"
  EditBox 105, 85, 60, 15, MAXIS_case_number
  EditBox 105, 105, 50, 15, CAF_datestamp
  DropListBox 105, 125, 140, 15, "Select One:"+chr(9)+"CAF (DHS-5223)"+chr(9)+"HUF (DHS-8107)"+chr(9)+"SNAP App for Srs (DHS-5223F)"+chr(9)+"ApplyMN"+chr(9)+"Combined AR for Certain Pops (DHS-3727)"+chr(9)+"CAF Addendum (DHS-5223C)", CAF_form
  CheckBox 110, 160, 30, 10, "CASH", CASH_on_CAF_checkbox
  CheckBox 150, 160, 35, 10, "SNAP", SNAP_on_CAF_checkbox
  CheckBox 190, 160, 35, 10, "EMER", EMER_on_CAF_checkbox
  ButtonGroup ButtonPressed
    OkButton 260, 310, 50, 15
    CancelButton 315, 310, 50, 15
    PushButton 125, 310, 15, 15, "!", tips_and_tricks_button
  Text 10, 10, 360, 10, "Start this script at the beginning of the interview and keep it running during the entire course of the interview."
  Text 10, 20, 60, 10, "This script will:"
  Text 20, 30, 170, 10, "- Guide you through all of the interview questions."
  Text 20, 40, 170, 10, "- Capture client answers for CASE:NOTE"
  Text 20, 50, 260, 10, "- Create a document of the interview answers to be saved in the ECF Case File."
  Text 20, 60, 245, 10, "- Provide verbiage guidance for consistent resident interview experience."
  Text 20, 70, 260, 10, "- Store the interview date, time, and legth in a database (an FNS requirement)."
  Text 50, 90, 50, 10, "Case number:"
  Text 10, 110, 90, 10, "Date Application Received:"
  Text 40, 130, 60, 10, "Actual CAF Form:"
  GroupBox 105, 145, 125, 30, "Programs marked on CAF"
  Text 145, 315, 105, 10, "Look for me for Tips and Tricks!"
  Text 20, 280, 315, 10, "How do you want to be alerted to updates needed to answers/information in following dialogs?"
  DropListBox 25, 290, 295, 45, "Alert at the time you attempt to save each page of the dialog."+chr(9)+"Alert only once completing and leaving the final dialog.", select_err_msg_handling
  GroupBox 10, 175, 355, 130, "How to interact with this Script"
  Text 20, 200, 335, 20, "The script will have a place to enter the answer from the CAF; a 'yes/no/blank' field plus an 'open' field to enter exactly what the CAF has listed on it."
  Text 30, 220, 305, 10, "Entering information in these fields should happen as you discuss this answer with the client."
  Text 30, 230, 315, 10, "The script will consider that question to have 'confirmed response' if these fields are completed."
  Text 20, 245, 340, 20, "Entering detail in 'Interview Notes' should happen for any information the client provides verbally upon discussion of that question. "
  Text 30, 265, 330, 10, "All detail should be entered in this field because it is important we are capturing the full conversation."
  Text 70, 185, 220, 10, "You should have this script running DURING the entire interview."
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


If select_err_msg_handling = "Alert at the time you attempt to save each page of the dialog." Then show_err_msg_during_movement = TRUE
If select_err_msg_handling = "Alert only once completing and leaving the final dialog." Then show_err_msg_during_movement = FALSE

show_known_addr = FALSE
vars_filled = FALSE

Call back_to_SELF
Call restore_your_work(vars_filled)			'looking for a 'restart' run
Call convert_date_into_MAXIS_footer_month(CAF_datestamp, MAXIS_footer_month, MAXIS_footer_year)
If vars_filled = TRUE Then show_known_addr = TRUE		'This is a setting for the address dialog to see the view

Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, unknown_cash_pending)

'If we already know the variables because we used 'restore your work' OR if there is no case number, we don't need to read the information from MAXIS
If vars_filled = FALSE AND no_case_number_checkbox = unchecked Then
	'Needs to determine MyDocs directory before proceeding.
	Set WshShell = CreateObject("WScript.Shell")

	user_myDocs_folder = WshShell.SpecialFolders("MyDocuments") & "\"
	intvw_msg_file = user_myDocs_folder & "interview message.txt"

	Set oExec = WshShell.Exec("notepad " & intvw_msg_file)

	Call back_to_SELF

	Call generate_client_list(all_the_clients, "Select or Type")				'Here we read for the clients and add it to a droplist
	list_for_array = right(all_the_clients, len(all_the_clients) - 15)			'Then we create an array of the the full hh list for looping purpoases
	full_hh_list = Split(list_for_array, chr(9))


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

		ReDim Preserve HH_MEMB_ARRAY(last_const, clt_count)
		HH_MEMB_ARRAY(ref_number, clt_count) = hh_clt
		' HH_MEMB_ARRAY(define_the_member, clt_count)

		Call navigate_to_MAXIS_screen("STAT", "MEMB")		'===============================================================================================
		EMWriteScreen HH_MEMB_ARRAY(ref_number, clt_count), 20, 76
		transmit

		EMReadScreen access_denied_check, 13, 24, 2         'Sometimes MEMB gets this access denied issue and we have to work around it.
		If access_denied_check = "ACCESS DENIED" Then
			PF10
			HH_MEMB_ARRAY(last_name_const, clt_count) = "UNABLE TO FIND"
			HH_MEMB_ARRAY(first_name_const, clt_count) = "Access Denied"
			HH_MEMB_ARRAY(mid_initial, clt_count) = ""
			HH_MEMB_ARRAY(access_denied, clt_count) = TRUE
		Else
			HH_MEMB_ARRAY(access_denied, clt_count) = FALSE
			EMReadscreen HH_MEMB_ARRAY(last_name_const, clt_count), 25, 6, 30
			EMReadscreen HH_MEMB_ARRAY(first_name_const, clt_count), 12, 6, 63
			EMReadscreen HH_MEMB_ARRAY(mid_initial, clt_count), 1, 6, 79
			EMReadScreen HH_MEMB_ARRAY(age, clt_count), 3, 8, 76

			EMReadScreen HH_MEMB_ARRAY(date_of_birth, clt_count), 10, 8, 42
			EMReadScreen HH_MEMB_ARRAY(ssn, clt_count), 11, 7, 42
			EMReadScreen HH_MEMB_ARRAY(ssn_verif, clt_count), 1, 7, 68
			EMReadScreen HH_MEMB_ARRAY(birthdate_verif, clt_count), 2, 8, 68
			EMReadScreen HH_MEMB_ARRAY(gender, clt_count), 1, 9, 42
			EMReadScreen HH_MEMB_ARRAY(race, clt_count), 30, 17, 42
			EMReadScreen HH_MEMB_ARRAY(spoken_lang, clt_count), 20, 12, 42
			EMReadScreen HH_MEMB_ARRAY(written_lang, clt_count), 29, 13, 42
			EMReadScreen HH_MEMB_ARRAY(interpreter, clt_count), 1, 14, 68
			EMReadScreen HH_MEMB_ARRAY(alias_yn, clt_count), 1, 15, 42
			EMReadScreen HH_MEMB_ARRAY(ethnicity_yn, clt_count), 1, 16, 68

			HH_MEMB_ARRAY(age, clt_count) = trim(HH_MEMB_ARRAY(age, clt_count))
			If HH_MEMB_ARRAY(age, clt_count) = "" Then HH_MEMB_ARRAY(age, clt_count) = 0
			HH_MEMB_ARRAY(age, clt_count) = HH_MEMB_ARRAY(age, clt_count) * 1
			HH_MEMB_ARRAY(last_name_const, clt_count) = trim(replace(HH_MEMB_ARRAY(last_name_const, clt_count), "_", ""))
			HH_MEMB_ARRAY(first_name_const, clt_count) = trim(replace(HH_MEMB_ARRAY(first_name_const, clt_count), "_", ""))
			HH_MEMB_ARRAY(mid_initial, clt_count) = replace(HH_MEMB_ARRAY(mid_initial, clt_count), "_", "")
			EMReadScreen HH_MEMB_ARRAY(id_verif, clt_count), 2, 9, 68

			EMReadScreen HH_MEMB_ARRAY(rel_to_applcnt, clt_count), 2, 10, 42              'reading the relationship from MEMB'
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "01" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Self"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "02" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Spouse"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "03" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Child"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "04" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Parent"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "05" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Sibling"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "06" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Step Sibling"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "08" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Step Child"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "09" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Step Parent"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "10" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Aunt"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "11" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Uncle"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "12" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Niece"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "13" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Nephew"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "14" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Cousin"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "15" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Grandparent"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "16" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Grandchild"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "17" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Other Relative"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "18" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Legal Guardian"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "24" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Not Related"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "25" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Live-in Attendant"
			If HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "27" Then HH_MEMB_ARRAY(rel_to_applcnt, clt_count) = "Unknown"

			If HH_MEMB_ARRAY(id_verif, clt_count) = "BC" Then HH_MEMB_ARRAY(id_verif, clt_count) = "BC - Birth Certificate"
			If HH_MEMB_ARRAY(id_verif, clt_count) = "RE" Then HH_MEMB_ARRAY(id_verif, clt_count) = "RE - Religious Record"
			If HH_MEMB_ARRAY(id_verif, clt_count) = "DL" Then HH_MEMB_ARRAY(id_verif, clt_count) = "DL - Drivers License/ST ID"
			If HH_MEMB_ARRAY(id_verif, clt_count) = "DV" Then HH_MEMB_ARRAY(id_verif, clt_count) = "DV - Divorce Decree"
			If HH_MEMB_ARRAY(id_verif, clt_count) = "AL" Then HH_MEMB_ARRAY(id_verif, clt_count) = "AL - Alien Card"
			If HH_MEMB_ARRAY(id_verif, clt_count) = "AD" Then HH_MEMB_ARRAY(id_verif, clt_count) = "AD - Arrival//Depart"
			If HH_MEMB_ARRAY(id_verif, clt_count) = "DR" Then HH_MEMB_ARRAY(id_verif, clt_count) = "DR - Doctor Stmt"
			If HH_MEMB_ARRAY(id_verif, clt_count) = "PV" Then HH_MEMB_ARRAY(id_verif, clt_count) = "PV - Passport/Visa"
			If HH_MEMB_ARRAY(id_verif, clt_count) = "OT" Then HH_MEMB_ARRAY(id_verif, clt_count) = "OT - Other Document"
			If HH_MEMB_ARRAY(id_verif, clt_count) = "NO" Then HH_MEMB_ARRAY(id_verif, clt_count) = "NO - No Veer Prvd"

			If HH_MEMB_ARRAY(age, clt_count) > 18 then
				HH_MEMB_ARRAY(cash_minor, clt_count) = FALSE
			Else
				HH_MEMB_ARRAY(cash_minor, clt_count) = TRUE
			End If
			If HH_MEMB_ARRAY(age, clt_count) > 21 then
				HH_MEMB_ARRAY(snap_minor, clt_count) = FALSE
			Else
				HH_MEMB_ARRAY(snap_minor, clt_count) = TRUE
			End If

			HH_MEMB_ARRAY(date_of_birth, clt_count) = replace(HH_MEMB_ARRAY(date_of_birth, clt_count), " ", "/")
			If HH_MEMB_ARRAY(birthdate_verif, clt_count) = "BC" Then HH_MEMB_ARRAY(birthdate_verif, clt_count) = "BC - Birth Certificate"
			If HH_MEMB_ARRAY(birthdate_verif, clt_count) = "RE" Then HH_MEMB_ARRAY(birthdate_verif, clt_count) = "RE - Religious Record"
			If HH_MEMB_ARRAY(birthdate_verif, clt_count) = "DL" Then HH_MEMB_ARRAY(birthdate_verif, clt_count) = "DL - Drivers License/State ID"
			If HH_MEMB_ARRAY(birthdate_verif, clt_count) = "DV" Then HH_MEMB_ARRAY(birthdate_verif, clt_count) = "DV - Divorce Decree"
			If HH_MEMB_ARRAY(birthdate_verif, clt_count) = "AL" Then HH_MEMB_ARRAY(birthdate_verif, clt_count) = "AL - Alien Card"
			If HH_MEMB_ARRAY(birthdate_verif, clt_count) = "DR" Then HH_MEMB_ARRAY(birthdate_verif, clt_count) = "DR - Doctor Statement"
			If HH_MEMB_ARRAY(birthdate_verif, clt_count) = "OT" Then HH_MEMB_ARRAY(birthdate_verif, clt_count) = "OT - Other Document"
			If HH_MEMB_ARRAY(birthdate_verif, clt_count) = "PV" Then HH_MEMB_ARRAY(birthdate_verif, clt_count) = "PV - Passport/Visa"
			If HH_MEMB_ARRAY(birthdate_verif, clt_count) = "NO" Then HH_MEMB_ARRAY(birthdate_verif, clt_count) = "NO - No Verif Provided"

			HH_MEMB_ARRAY(ssn, clt_count) = replace(HH_MEMB_ARRAY(ssn, clt_count), " ", "-")
			if HH_MEMB_ARRAY(ssn, clt_count) = "___-__-____" Then HH_MEMB_ARRAY(ssn, clt_count) = ""
			HH_MEMB_ARRAY(ssn_no_space, clt_count) = replace(HH_MEMB_ARRAY(ssn, clt_count), "-", "")

			If HH_MEMB_ARRAY(ssn_verif, clt_count) = "A" THen HH_MEMB_ARRAY(ssn_verif, clt_count) = "A - SSN Applied For"
			If HH_MEMB_ARRAY(ssn_verif, clt_count) = "P" THen HH_MEMB_ARRAY(ssn_verif, clt_count) = "P - SSN Provided, verif Pending"
			If HH_MEMB_ARRAY(ssn_verif, clt_count) = "N" THen HH_MEMB_ARRAY(ssn_verif, clt_count) = "N - SSN Not Provided"
			If HH_MEMB_ARRAY(ssn_verif, clt_count) = "V" THen HH_MEMB_ARRAY(ssn_verif, clt_count) = "V - SSN Verified via Interface"

			If HH_MEMB_ARRAY(gender, clt_count) = "M" Then HH_MEMB_ARRAY(gender, clt_count) = "Male"
			If HH_MEMB_ARRAY(gender, clt_count) = "F" Then HH_MEMB_ARRAY(gender, clt_count) = "Female"

			HH_MEMB_ARRAY(race, clt_count) = trim(HH_MEMB_ARRAY(race, clt_count))

			HH_MEMB_ARRAY(spoken_lang, clt_count) = replace(replace(HH_MEMB_ARRAY(spoken_lang, clt_count), "_", ""), "  ", " - ")
			HH_MEMB_ARRAY(written_lang, clt_count) = trim(replace(replace(replace(HH_MEMB_ARRAY(written_lang, clt_count), "_", ""), "  ", " - "), "(HRF)", ""))


			Call navigate_to_MAXIS_screen("STAT", "MEMI")		'===============================================================================================
			EMWriteScreen HH_MEMB_ARRAY(ref_number, clt_count), 20, 76
			transmit

			EMReadScreen HH_MEMB_ARRAY(marital_status, clt_count), 1, 7, 40
			EMReadScreen HH_MEMB_ARRAY(spouse_ref, clt_count), 2, 9, 49
			EMReadScreen HH_MEMB_ARRAY(spouse_name, clt_count), 40, 9, 52
			EMReadScreen HH_MEMB_ARRAY(last_grade_completed, clt_count), 2, 10, 49
			EMReadScreen HH_MEMB_ARRAY(citizen, clt_count), 1, 11, 49
			EMReadScreen HH_MEMB_ARRAY(other_st_FS_end_date, clt_count), 8, 13, 49
			EMReadScreen HH_MEMB_ARRAY(in_mn_12_mo, clt_count), 1, 14, 49
			EMReadScreen HH_MEMB_ARRAY(residence_verif, clt_count), 1, 14, 78
			EMReadScreen HH_MEMB_ARRAY(mn_entry_date, clt_count), 8, 15, 49
			EMReadScreen HH_MEMB_ARRAY(former_state, clt_count), 2, 15, 78

			If HH_MEMB_ARRAY(marital_status, clt_count) = "N" Then HH_MEMB_ARRAY(marital_status, clt_count) = "N - Never Married"
			If HH_MEMB_ARRAY(marital_status, clt_count) = "M" Then HH_MEMB_ARRAY(marital_status, clt_count) = "M - Married Living with Spouse"
			If HH_MEMB_ARRAY(marital_status, clt_count) = "S" Then HH_MEMB_ARRAY(marital_status, clt_count) = "S - Married Living Apart"
			If HH_MEMB_ARRAY(marital_status, clt_count) = "L" Then HH_MEMB_ARRAY(marital_status, clt_count) = "L - Legally Seperated"
			If HH_MEMB_ARRAY(marital_status, clt_count) = "D" Then HH_MEMB_ARRAY(marital_status, clt_count) = "D - Divorced"
			If HH_MEMB_ARRAY(marital_status, clt_count) = "W" Then HH_MEMB_ARRAY(marital_status, clt_count) = "W - Widowed"
			If HH_MEMB_ARRAY(spouse_ref, clt_count) = "__" Then HH_MEMB_ARRAY(spouse_ref, clt_count) = ""
			HH_MEMB_ARRAY(spouse_name, clt_count) = trim(HH_MEMB_ARRAY(spouse_name, clt_count))

			If HH_MEMB_ARRAY(last_grade_completed, clt_count) = "00" Then HH_MEMB_ARRAY(last_grade_completed, clt_count) = "Not Attended or Pre-Grade 1 - 00"
			If HH_MEMB_ARRAY(last_grade_completed, clt_count) = "12" Then HH_MEMB_ARRAY(last_grade_completed, clt_count) = "High School Diploma or GED - 12"
			If HH_MEMB_ARRAY(last_grade_completed, clt_count) = "13" Then HH_MEMB_ARRAY(last_grade_completed, clt_count) = "Some Post Sec Education - 13"
			If HH_MEMB_ARRAY(last_grade_completed, clt_count) = "14" Then HH_MEMB_ARRAY(last_grade_completed, clt_count) = "High School Plus Certiificate - 14"
			If HH_MEMB_ARRAY(last_grade_completed, clt_count) = "15" Then HH_MEMB_ARRAY(last_grade_completed, clt_count) = "Four Year Degree - 15"
			If HH_MEMB_ARRAY(last_grade_completed, clt_count) = "16" Then HH_MEMB_ARRAY(last_grade_completed, clt_count) = "Grad Degree - 16"
			If len(HH_MEMB_ARRAY(last_grade_completed, clt_count)) = 2 Then HH_MEMB_ARRAY(last_grade_completed, clt_count) = "Grade " & HH_MEMB_ARRAY(last_grade_completed, clt_count)
			If HH_MEMB_ARRAY(citizen, clt_count) = "Y" Then HH_MEMB_ARRAY(citizen, clt_count) = "Yes"
			If HH_MEMB_ARRAY(citizen, clt_count) = "N" Then HH_MEMB_ARRAY(citizen, clt_count) = "No"

			HH_MEMB_ARRAY(other_st_FS_end_date, clt_count) = replace(HH_MEMB_ARRAY(other_st_FS_end_date, clt_count), " ", "/")
			If HH_MEMB_ARRAY(other_st_FS_end_date, clt_count) = "__/__/__" Then HH_MEMB_ARRAY(other_st_FS_end_date, clt_count) = ""
			If HH_MEMB_ARRAY(in_mn_12_mo, clt_count) = "Y" Then HH_MEMB_ARRAY(in_mn_12_mo, clt_count) = "Yes"
			If HH_MEMB_ARRAY(in_mn_12_mo, clt_count) = "N" Then HH_MEMB_ARRAY(in_mn_12_mo, clt_count) = "No"
			If HH_MEMB_ARRAY(residence_verif, clt_count) = "1" Then HH_MEMB_ARRAY(residence_verif, clt_count) = "1 - Rent Receipt"
			If HH_MEMB_ARRAY(residence_verif, clt_count) = "2" Then HH_MEMB_ARRAY(residence_verif, clt_count) = "2 - Landlord's Statement"
			If HH_MEMB_ARRAY(residence_verif, clt_count) = "3" Then HH_MEMB_ARRAY(residence_verif, clt_count) = "3 - Utility Bill"
			If HH_MEMB_ARRAY(residence_verif, clt_count) = "4" Then HH_MEMB_ARRAY(residence_verif, clt_count) = "4 - Other"
			If HH_MEMB_ARRAY(residence_verif, clt_count) = "N" Then HH_MEMB_ARRAY(residence_verif, clt_count) = "N - Verif Not Provided"
			HH_MEMB_ARRAY(mn_entry_date, clt_count) = replace(HH_MEMB_ARRAY(mn_entry_date, clt_count), " ", "/")
			If HH_MEMB_ARRAY(mn_entry_date, clt_count) = "__/__/__" Then HH_MEMB_ARRAY(mn_entry_date, clt_count) = ""
			If HH_MEMB_ARRAY(former_state, clt_count) = "__" Then HH_MEMB_ARRAY(former_state, clt_count) = ""


		End If

		memb_droplist = memb_droplist+chr(9)+HH_MEMB_ARRAY(ref_number, clt_count) & " - " & HH_MEMB_ARRAY(full_name_const, clt_count)
		If HH_MEMB_ARRAY(fs_pwe, clt_count) = "Yes" Then the_pwe_for_this_case = HH_MEMB_ARRAY(ref_number, clt_count) & " - " & HH_MEMB_ARRAY(full_name_const, clt_count)

		' HH_MEMB_ARRAY(clt_count).intend_to_reside_in_mn = "Yes"

		' ReDim Preserve ALL_ANSWERS_ARRAY(ans_notes, clt_count)
		clt_count = clt_count + 1
	Next







	For the_members = 0 to UBound(HH_MEMB_ARRAY, 2)
		HH_MEMB_ARRAY(race_a_checkbox, the_members) = unchecked
		HH_MEMB_ARRAY(race_b_checkbox, the_members) = unchecked
		HH_MEMB_ARRAY(race_n_checkbox, the_members) = unchecked
		HH_MEMB_ARRAY(race_p_checkbox, the_members) = unchecked
		HH_MEMB_ARRAY(race_w_checkbox, the_members) = unchecked
		HH_MEMB_ARRAY(snap_req_checkbox, the_members) = unchecked
		If SNAP_on_CAF_checkbox = checked Then HH_MEMB_ARRAY(snap_req_checkbox, the_members) = checked
		HH_MEMB_ARRAY(cash_req_checkbox, the_members) = unchecked
		If CASH_on_CAF_checkbox = checked Then HH_MEMB_ARRAY(cash_req_checkbox, the_members) = checked
		HH_MEMB_ARRAY(emer_req_checkbox, the_members) = unchecked
		If EMER_on_CAF_checkbox = checked Then HH_MEMB_ARRAY(emer_req_checkbox, the_members) = checked
		HH_MEMB_ARRAY(none_req_checkbox, the_members) = unchecked

		HH_MEMB_ARRAY(clt_has_sponsor, the_members) = ""
		HH_MEMB_ARRAY(client_verification, the_members) = ""
		HH_MEMB_ARRAY(client_verification_details, the_members) = ""
		HH_MEMB_ARRAY(client_notes, the_members) = ""
	Next

	'Now we gather the address information that exists in MAXIS
	Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_addr_city, resi_addr_state, resi_addr_zip, resi_addr_county, addr_verif, homeless_yn, reservation_yn, living_situation, mail_line_one, mail_line_two, mail_addr_city, mail_addr_state, mail_addr_zip, addr_eff_date, addr_future_date, phone_one_number, phone_two_number, phone_three_number, phone_pne_type, phone_two_type, phone_four_type)
	resi_addr_street_full = resi_line_one & " " & resi_line_two
	resi_addr_street_full = trim(resi_addr_street_full)
	mail_addr_street_full = mail_line_one & " " & mail_line_two
	mail_addr_street_full = trim(mail_addr_street_full)

	arep_in_MAXIS = FALSE
	update_arep = TRUE
	Call access_AREP_panel(access_type, arep_name, arep_addr_street, arep_addr_city, arep_addr_state, arep_addr_zip, arep_phone_one, arep_ext_one, arep_phone_two, arep_ext_two, forms_to_arep, mmis_mail_to_arep)
	If arep_name <> "" Then
		arep_in_MAXIS = TRUE
		update_arep = FALSE
		arep_complete_forms_checkbox = checked
	End If
	If forms_to_arep = "Y" Then arep_get_notices_checkbox = checked

	show_known_addr = TRUE

	MsgBox "Press 'OK' when you have explained the interview to the resident."
	oExec.Terminate()
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
add_verif_13_btn			= 1076
add_job_btn					= 1077
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
add_verif_jobs_btn			= 1095
clear_job_btn				= 1100
open_r_and_r_button 		= 1200
caf_page_one_btn			= 1300
caf_addr_btn				= 1400
caf_membs_btn				= 1500
caf_q_1_6_btn				= 1600
caf_q_7_11_btn				= 1700
caf_q_12_13_btn				= 1800
caf_q_14_15_btn				= 1900
caf_q_16_20_btn				= 2000
caf_q_21_24_btn				= 2100
caf_qual_q_btn				= 2200
caf_last_page_btn			= 2300
finish_interview_btn		= 2400
exp_income_guidance_btn 	= 2500
return_btn 					= 900

btn_placeholder = 4000
for each_job = 0 to UBOUND(JOBS_ARRAY, 2)
	JOBS_ARRAY(jobs_edit_btn, each_job) = btn_placeholder
	btn_placeholder = btn_placeholder + 1
next
For btn_count = 0 to UBound(HH_MEMB_ARRAY, 2)
	HH_MEMB_ARRAY(button_one, btn_count) = 500 + btn_count
	HH_MEMB_ARRAY(button_two, btn_count) = 600 + btn_count
Next
interview_date = date
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

		Do
			' MsgBox page_display
			Dialog1 = Empty
			call define_main_dialog

			err_msg = ""

			prev_page = page_display


			dialog Dialog1
			cancel_confirmation
			' MsgBox  HH_MEMB_ARRAY(0).ans_imig_status
			save_your_work
			Call check_for_errors

			If show_err_msg_during_movement = FALSE AND ButtonPressed <> finish_interview_btn Then err_msg = ""

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
		Loop until err_msg = ""
		' MsgBox "ButtonPressed - " & ButtonPressed

		call dialog_movement


	Loop until leave_loop = TRUE
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

'Navigate back to self and to EDRS
Back_to_self
CALL navigate_to_MAXIS_screen("INFC", "EDRS")

For the_memb = 0 to UBound(HH_MEMB_ARRAY, 2)

	'Write in SSN number into EDRS
	EMwritescreen HH_MEMB_ARRAY(ssn_no_space, the_memb), 2, 7
	transmit
	Emreadscreen SSN_output, 7, 24, 2

	'Check to see what results you get from entering the SSN. If you get NO DISQ then check the person's name
	IF SSN_output = "NO DISQ" THEN
		EMWritescreen HH_MEMB_ARRAY(last_name_const, the_memb), 2, 24
		EMWritescreen HH_MEMB_ARRAY(first_name_const, the_memb), 2, 58
		EMWritescreen HH_MEMB_ARRAY(mid_initial, the_memb), 2, 76
		transmit
		EMreadscreen NAME_output, 7, 24, 2
		IF NAME_output = "NO DISQ" THEN        'If after entering a name you still get NO DISQ then let worker know otherwise let them know you found a name.
			HH_MEMB_ARRAY(edrs_msg, the_memb) = "No disqualifications found for Member #: " & HH_MEMB_ARRAY(ref_number, the_memb) & " " & HH_MEMB_ARRAY(first_name_const, the_memb) & " " & HH_MEMB_ARRAY(last_name_const, the_memb)
			HH_MEMB_ARRAY(edrs_match, the_memb) = FALSE
		ELSE
			HH_MEMB_ARRAY(edrs_msg, the_memb) = "Member #: " & HH_MEMB_ARRAY(ref_number, the_memb) & " " & HH_MEMB_ARRAY(first_name_const, the_memb) & " " & HH_MEMB_ARRAY(last_name_const, the_memb) & " has a potential name match."
			HH_MEMB_ARRAY(edrs_match, the_memb) = TRUE
		END IF
	ELSE
		HH_MEMB_ARRAY(edrs_msg, the_memb) = "Member #: " & HH_MEMB_ARRAY(ref_number, the_memb) & " " & HH_MEMB_ARRAY(first_name_const, the_memb) & " " & HH_MEMB_ARRAY(last_name_const, the_memb) & " has SSN Match."    'If after searching a SSN number you don't get the NO DISQ message then let worker know you found the SSN
		HH_MEMB_ARRAY(edrs_match, the_memb) = TRUE
	END IF
Next
Do
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 550, 385, "Full Interview Questions"
		  ButtonGroup ButtonPressed
		    PushButton 465, 365, 80, 15, "Continue", continue_btn
		    Text 10, 10, 320, 10, "EDRs has been completed for all Household Members."
			y_pos = 25
		    For the_memb = 0 to UBound(HH_MEMB_ARRAY, 2)
				Text 20, y_pos, 420, 10, HH_MEMB_ARRAY(edrs_msg, the_memb)

				PushButton 390, y_pos, 70, 10, "SSN SEARCH", HH_MEMB_ARRAY(button_one, the_memb)
				PushButton 460, y_pos, 70, 10, "NAME SEARCH", HH_MEMB_ARRAY(button_two, the_memb)
				If HH_MEMB_ARRAY(edrs_match, the_memb) = TRUE Then
					' GroupBox 15, y_pos - 15, 520, 50, "MEMB " & HH_MEMB_ARRAY(ref_number, the_memb) & " - " & HH_MEMB_ARRAY(full_name_const, the_memb)
					Text 30, y_pos + 20, 45, 10, "EDRs Notes:"
		  		    EditBox 80, y_pos + 15, 450, 15, HH_MEMB_ARRAY(edrs_notes, the_memb)
					y_pos = y_pos + 20
				End If
				' If HH_MEMB_ARRAY(edrs_match, the_memb) = FALSE Then GroupBox 15, y_pos - 15, 520, 30, "MEMB XX - MEMBER NAME"
				y_pos = y_pos + 20
			Next
		    Text 15, 350, 70, 10, "EDRs CASE Notes:"
		    EditBox 15, 360, 440, 15, edrs_notes_for_case
		EndDialog

		dialog Dialog1

		cancel_confirmation
		For the_memb = 0 to UBound(HH_MEMB_ARRAY, 2)
			If ButtonPressed = HH_MEMB_ARRAY(button_one, the_memb) OR ButtonPressed = HH_MEMB_ARRAY(button_two, the_memb) Then
				err_msg = err_msg & "LOOP"
				EMReadScreen edrs_check, 12, 1, 36
				If edrs_check <> "EDRS Inquiry" Then
					Back_to_self
					CALL navigate_to_MAXIS_screen("INFC", "EDRS")
				End If
				If ButtonPressed = HH_MEMB_ARRAY(button_two, the_memb) Then
					EMWritescreen HH_MEMB_ARRAY(last_name_const, the_memb), 2, 24
					EMWritescreen HH_MEMB_ARRAY(first_name_const, the_memb), 2, 58
					EMWritescreen HH_MEMB_ARRAY(mid_initial, the_memb), 2, 76
				End If
				If ButtonPressed = HH_MEMB_ARRAY(button_one, the_memb) Then EMwritescreen HH_MEMB_ARRAY(ssn_no_space, the_memb), 2, 7
				transmit
			End If
		Next

	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

call back_to_SELF

'CLIENT RESPONSIBILITEIS
Do
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 550, 385, "Full Interview Questions"
		  ButtonGroup ButtonPressed
		    PushButton 465, 365, 80, 15, "Continue", continue_btn
			PushButton 430, 22, 100, 13, "Open DHS 4163", open_r_and_r_btn
		  Text 10, 10, 160, 10, "REVIEW the information listedd here to the client:"
		  GroupBox 10, 25, 530, 335, "Rights and Responsibilities Text"
		  Text 20, 35, 505, 35, "Note: Cash on an Electronic Benefit Transfer (EBT) card is provided to help families meet their basic needs, including: food, shelter, clothing, utilities and transportation. These funds are provided until families can support themselves. It is illegal for an EBT user to buy or attempt to buy tobacco products or alcohol with the EBT card. If you do, it is fraud and you will be removed from the program. Do not use an EBT card at a gambling establishment or retail establishment, which provides adult-orientated entertainment in which performers disrobe or perform in an unclothed state for entertainment."
		  Text 20, 70, 275, 50, "- If you receive cash assistance and/or child care assistance, you must report changes which may affect your benefits to the county agency within 10 days after the change has occurred. If you receive Supplemental Nutrition Assistance Program (SNAP) benefits, report changes by the 10th of the month following the month of the change. Each program may have different requirements for reporting changes. Talk to your caseworker about what you must report."

		  Text 20, 120, 275, 10, "You may be required to report changes in:"
		  Text 20, 130, 275, 20, "-Employment - starting or stopping a job or business; change in hours, earnings or expenses"
		  Text 20, 150, 275, 25, "- Income - receipt or change in child support, Social Security, veteran benefits, unemployment insurance, inheritance or insurance benefits"
		  Text 20, 170, 275, 20, "- Property - purchase, sale or transfer of a house, car or other items of value, or if you receive an inheritance or settlement"
		  Text 20, 190, 275, 20, "- Household - When a person dies or becomes disabled, moves in or out of your home or temporarily leaves; pregnancy; birth of a child."
		  Text 20, 210, 275, 10, "- Citizenship or immigration status"
		  Text 20, 220, 275, 10, "- Address"
		  Text 20, 230, 275, 10, "- Housing costs and/or rent subsidy"
		  Text 20, 240, 275, 10, "- Utility costs"
		  Text 20, 250, 275, 10, "- Filing a lawsuit"
		  Text 20, 260, 275, 10, "- Absent parent custody or visits"
		  Text 20, 270, 275, 10, "- Drug felony conviction"
		  Text 20, 280, 275, 10, "- Marriage, separation or divorce"
		  Text 20, 290, 275, 10, "- School attendance"
		  Text 20, 300, 275, 10, "- Health insurance coverage and premiums"
		  Text 20, 315, 275, 20, "Note: If you change child care providers, you must tell your child care worker and provider at least 15 days before the change goes into effect."

		  Text 15, 335, 520, 10, "If you have any questions or are unsure about any reporting rules, contact your worker. If your worker is not available, leave a message so the worker can get back to you."

		  Text 310, 70, 225, 35, "- The county, state or federal agency may check any of the information you provide. To obtain some forms of information we must have your signed consent. If you don't allow the county to confirm your information, you might not receive assistance."
		  Text 310, 105, 225, 35, "- If you give us information you know is untrue, withhold information or do not report as required, or we discover your information is untrue, you may be investigated for fraud. This may result in you being disqualified from receiving benefits, charged criminally, or both."
		  Text 310, 140, 225, 50, "- The state or federal quality control agency may randomly choose your case for review. They will review statements you provided and will check to see if your eligibility was figured correctly. The state may seek information from other sources and will inform you about any contact they intend to make. If you do not cooperate, your benefits may stop."
		  Text 310, 195, 225, 10, "Cooperation requirements:"
		  Text 310, 205, 225, 45, "- If the county approves you for the Minnesota Family Investment Program (MFIP) or the Diversionary Work Program (DWP), you must cooperate with employment services, unless you are exempt. You must develop and sign an employment plan or your DWP application will be denied."
		  Text 310, 250, 225, 55, "- To receive MFIP, DWP, and/or child care assistance, you must cooperate with child support enforcement for all children in your household. You have the right to claim 'good cause' for not cooperating with child support enforcement. Yo must assign your child support to the state of Minnesota for all eligible children. If you do not cooperate or assign your child support, benefits will be denied or terminated."
		  Text 310, 305, 225, 30, "After the county approves your MFIP or DWP, if you receive child support directly from the noncustodial parent, you must report it to your worker."

		  Text 10, 370, 210, 10, "Confirm you have reviewed client responsibilities:"
		  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! Responsibilities Discussed"+chr(9)+"No, I could not complete this", confirm_resp_read
		EndDialog

		dialog Dialog1
		cancel_confirmation

		If ButtonPressed = open_r_and_r_btn Then
			err_msg = "LOOP"
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-4163-ENG"
		End If
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

'CLIENT RIGHTS
Do
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 550, 385, "Full Interview Questions"
		  ButtonGroup ButtonPressed
		    PushButton 465, 365, 80, 15, "Continue", continue_btn
			PushButton 430, 22, 100, 13, "Open DHS 4163", open_r_and_r_btn
		  Text 10, 10, 160, 10, "REVIEW the information listedd here to the client:"
		  GroupBox 10, 25, 530, 335, "Rights and Responsibilities Text"

		  Text 275, 35, 150, 10, "Your Rights"

		  Text 20, 50, 275, 30, "- Your right to privacy. Your private information, including your health information, is protected by state and federal laws. Your worker has given you a Notice of Privacy Practices (DHS-3979) information sheet explaining these rights."
		  Text 20, 85, 275, 10, "- You have the right to reapply at any time if your benefits stop."
		  Text 20, 95, 275, 20, "- You have the right to receive a paper OR electronic copy of your SNAP application."
		  Text 20, 105, 275, 25, "- You have the right to know why, if we have not processed your application within:"
		  Text 30, 115, 265, 20, "- 30 days for cash, SNAP and child care assistance"
		  Text 30, 125, 265, 20, "- 60 days for cash related to disability."
		  Text 20, 135, 275, 25, "- You have the right to know the rules of the program you are applying for and for the agency to tell you how your benefit amount was figured."
		  Text 20, 155, 275, 10, "- You have the right to choose where and with whom you live."
		  Text 20, 165, 275, 45, "- Expenses. You have the right to report expenses such as shelter, utilities, child care, child support or medical costs. These expenses may affect the amount of Supplemental Nutrition Assistance Program (SNAP) benefits that you receive. Failure to report or verify certain expenses listed will be a statement by your household that you do not want a deduction for the unreported expenses."

		  Text 310, 50, 225, 35, "For SNAP, you may appeal within 90 days by writing or calling the county or the State Appeals Office. You may represent yourself at the hearing, or you may have someone (an attorney, relative, friend or another person) speak for you."
		  Text 310, 90, 225, 50, "If you wish your assistance to continue until the hearing, you must appeal before the date of the proposed action or within 10 days after the date the agency notice was mailed, whichever is later. Ask your county or tribal worker to explain how the timing of your appeal could affect your present or future assistance."
		  Text 310, 140, 225, 20, "- Access to free legal services. Contact your worker for information on free legal services."
		  Text 310, 165, 225, 80, "- Appeal rights. If you are unhappy with the action taken or feel the agency did not act on your request for assistance, you may appeal. For cash, child care assistance and health care, you may appeal within 30 days from the date you receive the notice by writing to the county or tribal agency, or directly to the State Appeals Office at the Minnesota Department of Human Services, PO Box 64941, St. Paul, MN 55164-0941. (If you show good cause for not appealing your cash and health care within 30 days, the agency can accept your appeal for up to 90 days from the date you receive the notice.)"

		  Text 10, 370, 150, 10, "Confirm you have reviewed client rights:"
		  DropListBox 160, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! Rights Discussedd"+chr(9)+"No, I could not complete this", confirm_rights_read
		EndDialog

		dialog Dialog1

		If ButtonPressed = open_r_and_r_btn Then
			err_msg = "LOOP"
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-4163-ENG"
		End If

		cancel_confirmation
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

case_number_last_digit = right(MAXIS_case_number, 1)
case_number_last_digit = case_number_last_digit * 1
If case_number_last_digit = 4 Then snap_day_of_issuance = "4th"
If case_number_last_digit = 5 Then snap_day_of_issuance = "5th"
If case_number_last_digit = 6 Then snap_day_of_issuance = "6th"
If case_number_last_digit = 7 Then snap_day_of_issuance = "7th"
If case_number_last_digit = 8 Then snap_day_of_issuance = "8th"
If case_number_last_digit = 9 Then snap_day_of_issuance = "9th"
If case_number_last_digit = 0 Then snap_day_of_issuance = "10th"
If case_number_last_digit = 1 Then snap_day_of_issuance = "11th"
If case_number_last_digit = 2 Then snap_day_of_issuance = "12th"
If case_number_last_digit = 3 Then snap_day_of_issuance = "13th"
If case_number_last_digit MOD 2 = 1 Then cash_day_of_issuance = "2nd to last day"		'ODD Number
If case_number_last_digit MOD 2 = 0 Then cash_day_of_issuance = "last day"		'EVEN Number
If cash_type = "ADULT" Then cash_day_of_issuance = "first day"

'EBT RESPONSIBILITIES AND USAGE
Do
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 550, 385, "Full Interview Questions"
		  ButtonGroup ButtonPressed
		    PushButton 465, 365, 80, 15, "Continue", continue_btn
		  Text 10, 10, 160, 10, "REVIEW the information listedd here to the client:"
		  GroupBox 10, 25, 530, 335, "EBT Information"
		  Text 20, 35, 275, 10, "For Cash and Supplemental Nutrition Assistance Program (SNAP) benefits:"
		  Text 30, 45, 265, 25, "- Each time you use your Electronic Benefits Transfer (EBT) card or sign your check, you state that you have informed the county or tribal agency about any changes in your situation that may affect your benefits."
		  Text 30, 75, 265, 25, "- Each time your EBT card is used, we assume you have received your cash or SNAP benefits, unless you reported your card lost or stolen to the county or tribal agency."

		  Text 20, 105, 275, 25, "The standard way to get your benefits to you is through issuance on an EBT card. For cash benefits, there may be other options such as a vendor payment or direct deposit. If you want more information about these options, please let us know."

		  Text 20, 140, 275, 10, "EBT card balances and information can be found:"
		  Text 30, 150, 265, 10, "- Call customer service, 24 hours a day / 7 days a week - Toll-free: 888-997-2227"
		  Text 30, 160, 265, 25, "- Go to www.ebtEDGE.com - Under EBT Cardholders, click on 'More Information' and log in using your user ID and password."
		  ' Text 20, 105, 275, 25, ""

		  GroupBox 10, 190, 290, 75, "Your EBT Issuances"
		  Text 20, 205, 275, 10, "If approved, your SNAP benefits will regularly be issued on the " & snap_day_of_issuance & " of the month."
		  Text 20, 220, 275, 10, "If approved, your CASH benefits will regularly be issued on the " & cash_day_of_issuance & " of the month."
		  Text 20, 235, 275, 20, "*** Due to processing changes or delay in receipt of information issuances days may change, you should access EBT information directly to ensure benefits are available."


		  Text 310, 35, 225, 10, "Do you already have an EBT card for this case?"
		  ComboBox 310, 45, 225, 45, "Select or Type"+chr(9)+"Yes - I have my card."+chr(9)+"No - I used to but I've lost it."+chr(9)+"No - I never had a card for this case"+chr(9)+case_card_info, case_card_info

		  Text 310, 65, 225, 10, "Do you know how to use an EBT card?"
		  DropListBox 310, 75, 225, 45, "Select One..."+chr(9)+"Yes"+chr(9)+"No", clt_knows_how_to_use_ebt_card

		  Text 10, 370, 210, 10, "Confirm you have reviewed EBT Information:"
		  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! EBT Basics Discussed"+chr(9)+"No, I could not complete this", confirm_ebt_read
		EndDialog

		dialog Dialog1

		cancel_confirmation
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE


If clt_knows_how_to_use_ebt_card = "No" then

	Do
		Do
			err_msg = ""

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 550, 385, "Full Interview Questions"
			  ButtonGroup ButtonPressed
			    PushButton 465, 365, 80, 15, "Continue", continue_btn
			  Text 10, 5, 160, 10, "REVIEW the information listedd here to the client:"
			  GroupBox 10, 15, 530, 340, "How to Use Your Minnesota EBT Card"
			  Text 185, 25, 345, 10, "Your EBT card is a safe, convenient and easy way for you to get your cash and food benefits each month."
			  Text 10, 370, 210, 10, "Confirm you have reviewed How to Use EBT Information:"
			  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! EBT Detail Discussed"+chr(9)+"No, I could not complete this", confirm_ebt_how_to_read
			  ButtonGroup ButtonPressed
			    PushButton 440, 5, 100, 13, "Open DHS 3315A", open_ebt_brochure_btn
			  Text 20, 30, 65, 10, "How to get a card:"
			  Text 25, 40, 305, 10, "- Your first card will be mailed to you within 2 business days of your benefits being approved."
			  Text 25, 50, 130, 10, "- Replacement cards are also mailed."
			  Text 40, 60, 170, 10, "Call 1-888-997-2227 to request a replacement card"
			  Text 40, 70, 170, 10, "Cards take about 5 business days to arrive."
			  Text 40, 80, 275, 10, "There is a $2 charge for all replacement cards, which is reduced from your benefit."
			  Text 25, 90, 230, 20, "NOTE: If you have cash benefits, you will be issued a card that has your name on it. SNAP only cases to not have names on the EBT card."
			  Text 20, 115, 85, 10, "Where to use your card:"
			  Text 25, 125, 120, 10, "At a store 'point-of-sale' machine."
			  Text 25, 135, 75, 10, "At an ATM (Cash Only)"
			  Text 25, 145, 140, 10, "At a check cashing business (Cash Only)"
			  Text 365, 45, 80, 10, "Keep your card safe"
			  Text 375, 55, 120, 10, "Lost benefits will not be replaced."
			  Text 375, 65, 155, 15, "Do not leave your card lying around or lose it, treat it like a debit card or cash."
			  Text 365, 90, 110, 10, "Do not throw your card away"
			  Text 375, 105, 150, 20, "The same card will be used every month for as long as you have benefits."
			  Text 375, 130, 155, 20, "Even if your cases closes and reopens in the future the same card may be used."
			  Text 365, 155, 145, 10, "Misuse of your EBT Card is Unlawful"
			  Text 370, 170, 160, 20, "- Selling your card or PIN to others may result in criminal charges and your benefits may end."
			  Text 370, 190, 165, 20, "- Attempting to buy tobacco products or alcoholic beverages with your EBT Card is considered fraud."
			  Text 370, 210, 165, 20, "- Repeated loss of your card may cause a fraud investigation to be opened on you."
			  Text 20, 165, 105, 10, "How to get or change your PIN:"
			  Text 25, 180, 135, 10, "- Call customer service at 888-997-2227"
			  Text 25, 190, 165, 10, "- Visit your county or tribal human services office"
			  Text 25, 200, 195, 10, "- Visit the ebtEDGE cardholder portal www.ebtEDGE.com"
			  Text 25, 210, 195, 20, "- Access the ebtEDGE mobile application, www.FISGLOBal.COM/EBTEDGEMOBILE"
			  Text 20, 230, 145, 20, "4 failed attepts to enter your PIN will lock your card until 12:01 am the next day."
			  Text 20, 255, 185, 10, "Register to receive EBT Information by Text Message"
			  Text 35, 325, 135, 10, "- Current Balance (text 'BAL' to 42265)"
			  Text 35, 335, 145, 10, "- Last 5 transactions  (text 'MINI' to 42265)"
			  Text 25, 265, 135, 10, "1. Go to www.ebtEDGE.com and log in"
			  Text 25, 275, 80, 10, "2. Select 'EBT Account'"
			  Text 25, 285, 205, 10, "3. Select 'Messaging Registration' under the Account Services menu"
			  Text 25, 295, 140, 10, "4. Enter your mobile (cell) phone number."
			  Text 25, 305, 230, 10, "5. Check the box next to SMS Balance, then click the 'Update' button."
			  Text 25, 315, 190, 10, "6. Use the same mobil number and text for information:"
			EndDialog

			dialog Dialog1

			If ButtonPressed = open_ebt_brochure_btn Then
				err_msg = "LOOP"
				run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3315A-ENG"
			End If

			cancel_confirmation
		Loop until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = FALSE

End If


'NOTICE OF PRIVACY PRACTICES
Do
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 550, 385, "Full Interview Questions"
		  ButtonGroup ButtonPressed
		    PushButton 465, 365, 80, 15, "Continue", continue_btn
			PushButton 440, 5, 100, 13, "Open DHS 3979", open_npp_doc
		  Text 10, 5, 160, 10, "REVIEW the information listedd here to the client:"
		  GroupBox 10, 15, 530, 345, "Notice of Privacy Practices - About the Information you give us"
		  Text 20, 25, 505, 35, "This notice tells how private information about you may be used and disclosed and how you can get this information. Please review it carefully."
		  Text 15, 35, 275, 10, "Why do we ask for this information?"
		  Text 15, 45, 275, 10, "In order to determine whether and how we can help you, we collect information:"
		  Text 17, 55, 3, 10, "-"
		  Text 20, 55, 275, 10, "To tell you apart from other people with the same or similar name"
		  Text 17, 65, 3, 10, "-"
		  Text 20, 65, 275, 10, "To decide what you are eligible for"
		  Text 17, 75, 3, 10, "-"
		  Text 20, 75, 275, 20, "To help you get medical, mental health, financial or social services and decide if you can pay for some services"
		  Text 17, 95, 3, 10, "-"
		  Text 20, 95, 275, 10, "To decide if you or your family need protective services"
		  Text 17, 105, 3, 10, "-"
		  Text 20, 105, 275, 10, "To decide about out-of-home care and in-home care for you or your children"
		  Text 17, 115, 3, 10, "-"
		  Text 20, 115, 275, 10, "To investigate the accuracy of the information in your application"
		  Text 15, 125, 275, 20, "After we have begun to provide services or support to you, we may collect additional information:"
		  Text 17, 145, 3, 10, "-"
		  Text 20, 145, 275, 10, "To make reports, do research, do audits, and evaluate our programs"
		  Text 17, 155, 3, 10, "-"
		  Text 20, 155, 275, 10, "To investigate reports of people who may lie about the help they need"
		  Text 17, 165, 3, 10, "-"
		  Text 20, 165, 275, 20, "To collect money from other agencies, like insurance companies, if they should pay for your care"
		  Text 17, 180, 3, 10, "-"
		  Text 20, 180, 275, 10, "To collect money from the state or federal government for help we give you."
		  Text 17, 190, 3, 10, "-"
		  Text 20, 190, 275, 20, "When your or your family's circumstances change and you are required to report the change (see Client Responsibilities and Rights - DHS-4163)"
		  Text 15, 210, 275, 10, "Why do we ask you for your Social Security number?"
		  Text 20, 220, 275, 75, "We need your Social Security number to give you medical assistance, some kinds of financial help, or child support enforcement services (42 CFR 435.910 [2006]; Minn. Stat. 256D.03, subd.3(h); Minn. Stat.256L.04, subd. 1a; 45 CFR 205.52 [2001]; 42 USC 666; 45 CFR 303.30 [2001]). We also need your Social Security Number to verify identity and prevent duplication of state and federal benefits. Additionally, your Social Security Number is used to conduct computer data matches with collaborative, nonprofit and private agencies to verify income, resources, or other information that may affect your eligibility and/or benefits."
		  Text 20, 285, 275, 10, "You do not have to give us the Social Security Number:"
		  Text 22, 295, 3, 10, "-"
		  Text 25, 295, 275, 10, "For persons in your home who are not applying for coverage"
		  Text 22, 305, 3, 10, "-"
		  Text 25, 305, 275, 10, "If you have religious objections"
		  Text 22, 315, 3, 10, "-"
		  Text 25, 315, 500, 10, "If you are not a United States citizen and are applying for Emergency Medical Assistance only"
		  Text 22, 325, 3, 10, "-"
		  Text 25, 325, 500, 20, "If you are from another country, in the United States on a temporary basis and do not have permission from the United States Citizenship and Immigration Services to live in the United States permanently"
		  Text 22, 342, 3, 10, "-"
		  Text 25, 342, 500, 10, "If you are living in the United States without the knowledge or approval of the U.S. Citizenship and Immigration Services."
		  Text 305, 35, 225, 10, "Do you have to answer the questions we ask?"
		  Text 310, 45, 240, 45, "You do not have to give us your personal information. Without the information, we may not be able to help you. If you give us wrong information on purpose, you can be investigated and charged with fraud."
		  Text 305, 75, 225, 10, "With whom may we share information?"
		  Text 305, 85, 225, 35, "We will only share information about you as needed and as allowed or required by law. We may share your information with the following agencies or persons who need the information to do their jobs:"
		  Text 307, 110, 3, 10, "-"
		  Text 310, 110, 225, 35, "Employees or volunteers with other state, county, local, federal, collaborative, nonprofit and private agencies"
		  Text 307, 130, 3, 10, "-"
		  Text 310, 130, 225, 35, "Researchers, auditors, investigators, and others who do quality of care reviews and studies or commence prosecutions or legal actions related to managing the human services programs."
		  Text 307, 155, 3, 10, "-"
		  Text 310, 155, 225, 35, "Court officials, county attorney, attorney general, other law enforcement officials, child support officials, and child protection and fraud investigators"
		  Text 307, 180, 3, 10, "-"
		  Text 310, 180, 225, 10, "Human services offices, including child support enforcement offices"
		  Text 307, 190, 3, 10, "-"
		  Text 310, 190, 225, 20, "Governmental agencies in other states administering public benefits programs"
		  Text 307, 210, 3, 10, "-"
		  Text 310, 210, 225, 20, "Health care providers, including mental health agencies and drug and alcohol treatment facilities"
		  Text 307, 230, 3, 10, "-"
		  Text 310, 230, 225, 20, "Health care insurers, health care agencies, managed care organizations and others who pay for your care"
		  Text 307, 250, 3, 10, "-"
		  Text 310, 250, 225, 10, "Guardians, conservators or persons with power of attorney"
		  Text 307, 260, 3, 10, "-"
		  Text 310, 260, 225, 20, "Coroners and medical investigators if you die and they investigate your death"
		  Text 307, 280, 3, 10, "-"
		  Text 310, 280, 225, 20, "Credit bureaus, creditors or collection agencies if you do not pay fees you owe to us for services"
		  Text 307, 300, 3, 10, "-"
		  Text 310, 300, 225, 10, "Anyone else to whom the law says we must or can give the information"
		  Text 10, 370, 210, 10, "Confirm you have reviewed Privacy Practices Information:"
		  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! Notice of Privacy Information Discussed"+chr(9)+"No, I could not complete this", confirm_npp_info_read
		EndDialog

		dialog Dialog1

		If ButtonPressed = open_npp_doc Then
			err_msg = "LOOP"
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3979-ENG"
		End If

		cancel_confirmation
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

Do
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 550, 385, "Full Interview Questions"
		  ButtonGroup ButtonPressed
		    PushButton 465, 365, 80, 15, "Continue", continue_btn
			PushButton 440, 5, 100, 13, "Open DHS 3979", open_npp_doc
		  Text 10, 5, 160, 10, "REVIEW the information listedd here to the client:"
		  GroupBox 10, 15, 530, 345, "Notice of Privacy Practices - Rights"
		  Text 20, 25, 505, 35, "This notice tells how private information about you may be used and disclosed and how you can get this information. Please review it carefully."
		  Text 15, 40, 275, 10, "What are your rights regarding the information we have about you?"
		  Text 17, 50, 3, 10, "-"
		  Text 20, 50, 275, 20, "You and people you have given permission to may see and copy private information we have about you. You may have to pay for the copies."
		  Text 17, 70, 3, 10, "-"
		  Text 20, 70, 275, 40, "You may question if the information we have about you is correct. Send your concerns in writing. Tell us why the information is wrong or not complete. Send your own explanation of the information you do not agree with. We will attach your explanation any time information is shared with another agency."
		  Text 17, 110, 3, 10, "-"
		  Text 20, 110, 275, 35, "You have the right to ask us in writing to share information with you in a certain way or in a certain place. For example, you may ask us to send health information to your work address instead of your home address. If we find that your request is reasonable, we will grant it."
		  Text 17, 150, 3, 10, "-"
		  Text 20, 150, 275, 20, "You have the right to ask us to limit or restrict the way that we use or disclose your information, but we are not required to agree to this request."
		  Text 17, 170, 3, 10, "-"
		  Text 20, 170, 275, 20, "If you do not understand the information, ask your worker to explain it to you. You can ask the Minnesota Department of Human Services for another copy of this notice."

		  Text 15, 200, 150, 10, "What privacy rights do children have?"
		  Text 20, 215, 490, 50, "If you are under 18, when parental consent for medical treatment is not required, information will not be shown to parents unless the health care provider believes not sharing the information would risk your health. Parents may see other information about you and let others see this information, unless you have asked that this information not be shared with your parents. You must ask for this in writing and say what information you do not want to share and why. If the agency agrees that sharing the information is not in your best interest, the information will not be shared with your parents. If the agency does not agree, the information may be shared with your parents if they ask for it."
		  Text 15, 270, 275, 10, "What if you believe your privacy rights have been violated?"
		  Text 20, 285, 490, 20, "If you think that the Minnesota Department of Human Services has violated your privacy rights, you may send a written complaint to the U.S. Department of Health and Human Services to the address below:"
		  Text 20, 305, 275, 10, "Minnesota Department of Human Services"
		  Text 20, 315, 275, 10, "Attn: Privacy Official"
		  Text 20, 325, 275, 10, "PO Box 64998"
		  Text 20, 335, 275, 10, "St. Paul, MN 55164-0998"

		  Text 305, 40, 225, 10, "What are our responsibilities?"
		  Text 307, 50, 3, 10, "-"
		  Text 310, 50, 225, 20, "We must protect the privacy of your private information according to the terms of this notice."
		  Text 307, 70, 3, 10, "-"
		  Text 310, 70, 225, 40, "We may not use your information for reasons other than the reasons listed on this form or share your information with individuals and agencies other than those listed on this form unless you tell us in writing that we can."
		  Text 307, 110, 3, 10, "-"
		  Text 310, 110, 225, 40, "We must follow the terms of this notice, but we may change our privacy policy because privacy laws change. We will put changes to our privacy rules on our website at: http://edocs.dhs.state.mn.us/lfserver/Public/DHS-3979-ENG"
		  ' ButtonGroup ButtonPressed
		  '   PushButton 310, 150, 100, 13, "Open DHS 3979", open_npp_doc
		  Text 10, 370, 210, 10, "Confirm you have reviewed Privacy Practices Rights:"
		  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! Notice of Privacy Rights Discussed"+chr(9)+"No, I could not complete this", confirm_npp_rights_read
		EndDialog

		dialog Dialog1

		If ButtonPressed = open_npp_doc Then
			err_msg = "LOOP"
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-3979-ENG"
		End If

		cancel_confirmation
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

'NOTICE ABOUT IEVS
Do
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 550, 385, "Full Interview Questions"
		  ButtonGroup ButtonPressed
		    PushButton 465, 365, 80, 15, "Continue", continue_btn
		  Text 10, 5, 160, 10, "REVIEW the information listedd here to the client:"
		  GroupBox 10, 15, 530, 345, "IEVS Information"
		  Text 15, 25, 275, 10, "What is the Income and Eligibility Verification System (IEVS)?"
		  Text 20, 35, 275, 20, "The government has a way to check income. It is the 'Income and Eligibility Verification System' (IEVS)."
		  Text 20, 55, 275, 30, "The law has us check your income with other agencies. We have to check income for all who ask for or get cash assistance, Supplemental Nutrition Assistance Program (SNAP) benefits or Medical Assistance (MA). This includes your children."
		  Text 20, 85, 275, 30, "We need Social Security Numbers (SSN) for anyone wanting help. If you have no SSN, you must apply for one. Apply with your county human services agency. You must report all SSNs to your worker."

		  Text 15, 115, 275, 10, "Agencies we get information from. We must trade facts with these agencies:"
		  Text 17, 125, 3, 10, "-"
		  Text 20, 125, 275, 10, "United States Social Security Administration (SSA)"
		  Text 17, 135, 3, 10, "-"
		  Text 20, 135, 275, 10, "United States Internal Revenue Service (IRS)"
		  Text 17, 145, 3, 10, "-"
		  Text 20, 145, 275, 10, "Minnesota Department of Employment and Economic Development (DEED)"
		  Text 17, 155, 3, 10, "-"
		  Text 20, 155, 275, 10, "Minnesota Office of Child Support Division"
		  Text 17, 165, 3, 10, "-"
		  Text 20, 165, 275, 10, "Agencies in other states that manage:"
		  Text 17, 175, 3, 10, "-"
		  Text 20, 175, 275, 10, "Unemployment Insurance"
		  Text 17, 185, 3, 10, "-"
		  Text 20, 185, 275, 10, "Cash assistance/SNAP/MA"
		  Text 17, 195, 3, 10, "-"
		  Text 20, 195, 275, 10, "Child support"
		  Text 17, 205, 3, 10, "-"
		  Text 20, 205, 275, 10, "SSI state supplements"
		  Text 15, 215, 275, 30, "These agencies have the right to get certain facts from us about you. They have to use those facts for programs like RSDI, child support, cash assistance, SNAP, MA, Unemployment Insurance, and SSI."

		  Text 15, 230, 275, 10, "Your duty to report"
		  Text 20, 240, 275, 10, "You must report all of your income and assets."
		  Text 20, 250, 275, 20, "You must still report all of your income, assets and other information on redetermination forms we send you.  "
		  Text 20, 270, 275, 20, "You must help the county agency check your income, assets and health insurance. IEVS is one way of proving your income, assets and health insurance amounts."
		  Text 15, 290, 275, 10, "What if you do not help"
		  Text 20, 300, 275, 20, "You must help us check your income, assets and health insurance to get cash assistance, SNAP and MA. If you don't, you and your family will not get help."

		  Text 120, 330, 380, 20, "Legal Authority - IEVS - 7 CFR, parts 271, 272, 273, 275; 42 CFR, parts 431, 435; 45 CFR, parts 205, 206, 233 - Work Reporting - Minnesota Statutes Section 256.998, Subd. 10"

		  Text 305, 25, 225, 10, "What facts will we get? How will we use them?"
		  Text 305, 35, 225, 40, "We check with other agencies about your income, assets and health insurance. If you didn't tell us about all of your income or assets, we will refigure your aid. Your aid might go lower or stop. If you get aid you should not be getting, we may use these facts in civil or criminal lawsuits."
		  Text 305, 75, 225, 40, "We will tell you if facts from other agencies are not the same as the facts you gave us. We will tell you what facts we got, the kind of income or assets, and the amount. We give you 10 days to respond in writing to prove if our facts are wrong."
		  Text 305, 115, 225, 40, "We will ask you to show proof of income, assets, or health insurance you did not report or that we could not verify. You may need to give us permission to check the facts with the source of data. We will tell you what happens if you do not sign for permission or do not help us."

		  Text 305, 155, 225, 10, "What is the Work Reporting System?"
		  Text 305, 165, 225, 40, "Minnesota employers must tell us when they hire someone. This information is used by the Child Support Program. We also use this information to see if a new employee is getting help from any of the programs listed above."
		  Text 305, 205, 225, 10, "How do we use it?"
		  Text 305, 215, 225, 40, "If the employee is getting help from any of these programs, the county worker gets a notice. If the client did not report the new job, the county worker will contact the client. The county worker may ask the client to show proof about the job. The client may need to give the county permission to check the facts with the employer. If a client does not help us check the information, they will lose benefits."
		  Text 305, 265, 225, 10, "The law limits who gets facts about you"
		  Text 305, 275, 225, 50, "The law limits the facts about you that we get from other agencies and the facts we give them. Contracts with the Minnesota Department of Human Services and those agencies also protect you. Only those agencies, the state, and the county agency where you apply for and get program benefits can use the facts about you. No one else can get the facts about you without your written permission."
		  ButtonGroup ButtonPressed
		    PushButton 15, 330, 100, 13, "Open DHS 2759", open_IEVS_doc
		  Text 10, 370, 210, 10, "Confirm you have reviewed IEVS Information:"
		  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! IEVS Information Discussed"+chr(9)+"No, I could not complete this", confirm_npp_rights_read
		EndDialog

		dialog Dialog1

		If ButtonPressed = open_IEVS_doc Then
			err_msg = "LOOP"
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-2759-ENG"
		End If

		cancel_confirmation
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

'APPEAL RIGHTS
Do
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 550, 385, "Full Interview Questions"
		  ButtonGroup ButtonPressed
		    PushButton 465, 365, 80, 15, "Continue", continue_btn
		  Text 10, 5, 160, 10, "REVIEW the information listedd here to the client:"
		  GroupBox 10, 15, 530, 345, "Appeal Rights"
		  Text 15, 25, 505, 10, "Appeal rights. An appeal is a legal process where a human services judge reviews a decision made by the agency. You may appeal a decision if:"
		  Text 20, 35, 500, 10, "You feel the agency did not act on your request for assistance."
		  Text 20, 45, 500, 10, "You do not agree with the action taken."
		  Text 15, 55, 505, 10, "You may represent yourself at the hearing, or you may have someone (an attorney, relative, friend or another person) speak for you."

		  Text 20, 65, 500, 20, "For emergency help, when your case is about an emergency and you need a faster decision on your appeal, you can ask for an emergency hearing in your appeal request. You can also request it by calling the Department of Human Services Appeals Division."
		  Text 20, 85, 500, 40, "For cash, child care and health care, you may appeal within 30 days from the date you received this notice by sending a written appeal request saying you do not agree with the decision. You can send this letter to the agency, or directly to the Appeals Division. If you show good cause for not appealing your cash, child care and health care within 30 days, the agency can accept your appeal for up to 90 days from the date of the notice. Good cause is when you have a good reason for not appealing on time. The Appeals Division will decide if your reason is a good cause reason. You can ask to meet informally with agency staff to try to solve the problem, but this meeting will not delay or replace your right to an appeal."
		  Text 20, 125, 500, 10, "For the Supplemental Nutrition Assistance Program, you may appeal within 90 days by writing or calling the agency or the Appeals Division."
		  Text 20, 135, 500, 10, "Submit your appeal request:"
		  Text 25, 145, 495, 10, "Online: https://edocs.dhs.state.mn.us/lfserver/Public/DHS-0033-ENG"
		  Text 25, 155, 495, 10, "Write: Minnesota Department of Human Services Appeals Division P.O. Box 64941 St. Paul, MN 55164-0941"
		  Text 25, 165, 495, 10, "Fax: 651-431-7523"
		  Text 25, 175, 495, 10, "Call: Metro: 651-431-3600 Greater Minnesota: 800-657-3510 "
		  Text 20, 185, 500, 40, "If you want to keep receiving your benefits until the hearing, you must appeal within 10 days of the date on the agency’s notice of action letter or before the proposed action takes place in order to keep benefits in place. For most programs, if you file your appeal on time, you will get your benefits until the Appeals Division decides your appeal. If you lose your appeal, you may have to pay back the benefits you got while your appeal was pending. You can ask the agency to end your benefits until the decision. If you end your benefits and then win your appeal, you will be paid back for benefits that you should have received or, for child care assistance, your provider will be reimbursed for eligible costs that you paid or incurred. Ask your agency worker to explain how the timing of your appeal could affect your present or future assistance."
		  Text 15, 235, 505, 10, "You have the right to reapply at any time if your benefits stop."
		  Text 15, 245, 505, 20, "Access to free legal services. You may be able to get legal advice or help with an appeal from your local legal aid office. To find your local legal aid office, visit www.LawHelpMN.org or call 888-354-5522."

		  ButtonGroup ButtonPressed
		    PushButton 15, 265, 100, 13, "Open DHS 3353", open_appeal_rights_doc
		  Text 10, 370, 210, 10, "Confirm you have reviewed Appeal Rights:"
		  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! Appeal Rights Discussed"+chr(9)+"No, I could not complete this", confirm_appeal_rights_read
		EndDialog

		dialog Dialog1

		If ButtonPressed = open_appeal_rights_doc Then
			err_msg = "LOOP"
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-3353-ENG"
		End If

		cancel_confirmation
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

'CIVIL RIGHTS NOTICE AND COMPLAINTS
Do
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 550, 385, "Full Interview Questions"
		  ButtonGroup ButtonPressed
		    PushButton 465, 365, 80, 15, "Continue", continue_btn
		  Text 10, 5, 160, 10, "REVIEW the information listedd here to the client:"
		  GroupBox 10, 15, 530, 345, "Civil Rights Notice and Complaints"
		  Text 15, 25, 505, 10, "Discrimination is against the law. The Minnesota Department of Human Services (DHS) does not discriminate on the basis of any of the following:"
		  Text 20, 35, 505, 10, "- race   - national origin   - religion   - public assistance status   - age   - sex   - color   - creed   - sexual orientation   - marital status   - disability   - political beliefs"

		  Text 15, 50, 275, 10, "Civil Rights Complaints"
		  Text 20, 60, 275, 20, "You have the right to file a discrimination complaint if you believe you were treated in a discriminatory way by a human services agency."
		  Text 20, 80, 275, 10, "Contact DHS directly only if you have a discrimination complaint:"
		  Text 25, 90, 275, 10, "Civil Rights Coordinator"
		  Text 25, 100, 275, 10, "Minnesota Department of Human Services"
		  Text 25, 110, 275, 10, "Equal Opportunity and Access Division"
		  Text 25, 120, 275, 10, "P.O. Box 64997 St. Paul, MN 55164-0997"
		  Text 25, 130, 275, 10, "651-431-3040 (voice) or use your preferred relay service"

		  Text 15, 140, 275, 10, "Minnesota Department of Human Rights (MDHR)"
		  Text 20, 150, 275, 20, "In Minnesota, you have the right to file a complaint with the MDHR if you believe you have been discriminated against because of any of the following:"
		  Text 25, 170, 275, 10, "- race   - sex   - color   - sexual orientation   - national origin   - marital status"
		  Text 25, 180, 275, 10, "- religion   - public assistance status   - creed   - disability"
		  Text 20, 190, 275, 10, "Contact the MDHR directly to file a complaint:"
		  Text 25, 200, 275, 10, "Minnesota Department of Human Rights"
		  Text 25, 210, 275, 10, "Freeman Building, 625 North Robert Street St. Paul, MN 55155"
		  Text 25, 220, 275, 10, "651-539-1100 (voice) 1-800-657-3704 (toll free) 651-296-9042 (fax)"
		  Text 25, 230, 275, 10, "Info.MDHR@state.mn.us (email)"


		  Text 15, 240, 275, 10, "U.S. Department of Health and Human Services' Office for Civil Rights (OCR)"
		  Text 20, 250, 275, 20, "You have the right to file a complaint with the OCR, a federal agency, if you believe you have been discriminated against because of any of the following:"
		  Text 25, 270, 275, 10, "- race   - age   - religion   - color   - disability   - national origin   - sex"
		  Text 20, 280, 275, 10, "Contact the OCR directly to file a complaint:"
		  Text 25, 290, 275, 10, "Director, U.S. Department of Health and Human Services' Office for Civil Rights"
		  Text 25, 300, 275, 10, "200 Independence Avenue SW, Room 509F HHH Building Washington, DC 20201"
		  Text 25, 310, 275, 10, "1-800-368-1019 (voice)  1-800-537-7697 (TDD)"
		  Text 25, 320, 275, 10, "Complaint Portal: https://ocrportal.hhs.gov/ocr/portal/lobby.jsf"

		  Text 305, 55, 225, 60, "In accordance with Federal civil rights law and U.S. Department of Agriculture (USDA) civil rights regulations and policies, the USDA, its Agencies, offices, and employees, and institutions participating in or administering USDA programs are prohibited from discriminating based on race, color, national origin, sex, religious creed, disability, age, political beliefs, or reprisal or retaliation for prior civil rights activity in any program or activity conducted or funded by USDA."
		  Text 305, 115, 225, 70, "Persons with disabilities who require alternative means of communication for program information (e.g. Braille, large print, audiotape, American Sign Language, etc.), should contact the Agency (State or local) where they applied for benefits. Individuals who are deaf, hard of hearing or have speech disabilities may contact USDA through the Federal Relay Service at 1-800-877-8339. Additionally, program information may be made available in languages other than English."
		  Text 305, 185, 225, 60, "To file a program complaint of discrimination, complete the USDA Program Discrimination Complaint Form, (AD-3027) found online at: http://www.ascr.usda.gov/complaint_filing_cust.html, and at any USDA office, or write a letter addressed to USDA and provide in the letter all of the information requested in the form. To request a copy of the complaint form, call 1-866- 632-9992. Submit your completed form or letter to USDA by:"
		  Text 310, 245, 225, 10, "(1) mail: U.S. Department of Agriculture"
		  Text 315, 255, 225, 10, "Office of the Assistant Secretary for Civil Rights"
		  Text 315, 265, 225, 10, "1400 Independence Avenue, SW"
		  Text 315, 275, 225, 10, "Washington, DC 20250-9410;"
		  Text 310, 285, 225, 10, "(2) fax: 202-690-7442; or"
		  Text 310, 295, 225, 10, "(3) email: program.intake@usda.gov"
		  Text 310, 305, 225, 10, "This institution is an equal opportunity provider."

		  ButtonGroup ButtonPressed
		    PushButton 15, 340, 100, 13, "Open DHS 3353", open_civil_rights_rights_doc
		  Text 10, 370, 210, 10, "Confirm you have reviewed Civil Rights Information:"
		  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! Civil Rights Discussed"+chr(9)+"No, I could not complete this", confirm_civil_rights_read
		EndDialog

		dialog Dialog1

		If ButtonPressed = open_civil_rights_rights_doc Then
			err_msg = "LOOP"
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-3353-ENG"
		End If

		cancel_confirmation
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

'COVER LETTER
Do
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 550, 385, "Full Interview Questions"
		  ButtonGroup ButtonPressed
		    PushButton 465, 365, 80, 15, "Continue", continue_btn
		  Text 10, 5, 160, 10, "REVIEW the information listedd here to the client:"
		  GroupBox 10, 15, 530, 345, "Hennepin County Cover Letter"



		  Text 10, 370, 210, 10, "Confirm you have reviewed Hennepin County Information Information:"
		  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! Cover Letter Discussed"+chr(9)+"No, I could not complete this", confirm_cover_letter_read
		EndDialog

		dialog Dialog1


		cancel_confirmation
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

'PROGRAM INFORMATION FOR CASH, FOOD, CHILD CARE - 2920
Do
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 550, 385, "Full Interview Questions"
		  ButtonGroup ButtonPressed
		    PushButton 465, 365, 80, 15, "Continue", continue_btn
		  Text 10, 5, 160, 10, "REVIEW the information listedd here to the client:"
		  GroupBox 10, 15, 530, 345, "Program Information for cash, food, and child care programs"
		  Text 15, 25, 505, 10, "How do you apply for help?"
		  Text 20, 35, 505, 10, "If you do not have enough money to meet your basic needs, you can apply to find out if you are eligible for these assistance programs."
		  Text 25, 45, 505, 10, "Apply online at MNBenefits.org or www.applymn.dhs.mn.gov"
		  Text 25, 55, 505, 10, "Mail or bring your completed application to your county human services agency"
		  Text 20, 65, 505, 10, "Food and cash programs require an interview with a worker. Most of the time this can be a phone interview. You will need to bring proof of:"
		  Text 20, 75, 505, 10, "- Who you are   - Where you live   - What family members live with you   - What your income is   - What you own."
		  Text 20, 85, 505, 10, "Whether or not you can receive help and how much you receive may depend on:"
		  Text 20, 95, 505, 10, "- How long you have lived in Minnesota   - How many people live with you   - How much income you and these people receive each month."
		  Text 20, 105, 505, 10, "Each program has different rules."

		  Text 15, 115, 275, 20, "Cash assistance is provided to help you meet your basic needs, if you are eligible. Some of the programs have time limits. Cash programs include:"
		  Text 20, 135, 275, 10, "Diversionary Work Program (DWP)"
		  Text 25, 145, 275, 30, "A short-term work program that provides employment services and basic living costs to eligible families. DWP is for families who are working or looking for work, but need help with basic living expenses and have not MFIP or DWP in the last 12 months."
		  Text 20, 175, 275, 10, "Minnesota Family Investment Program (MFIP)"
		  Text 25, 185, 275, 20, "A monthly cash assistance program for families with children under 19 or pregnant women, and who have low incomes."
		  Text 20, 205, 275, 10, "General Assistance (GA)"
		  Text 25, 215, 275, 10, "A monthly cash payment for adults who are unable to work who:"
		  Text 30, 225, 275, 10, "- Have little or no income and will soon return to work, or"
		  Text 30, 235, 275, 10, "- Are waiting to get help from other state or federal programs."
		  Text 20, 245, 275, 10, "Minnesota Supplemental Aid (MSA)"
		  Text 25, 255, 275, 10, "A small extra monthly cash payment for adults who are eligible for federal SSI."
		  Text 20, 265, 275, 10, "Group Residential Housing (GRH)"
		  Text 25, 275, 275, 20, "A monthly payment that helps pay room and board costs for people who live in authorized settings and are:"
		  Text 130, 285, 275, 10, "- Age 65 or older "
		  Text 130, 295, 275, 10, "- Disabled and age 18 or older, or "
		  Text 130, 305, 275, 10, "- Have blindness."
		  Text 20, 315, 275, 10, "Refugee Cash Assistance (RCA)"
		  Text 25, 325, 275, 10, "A monthly cash payment for refugees and asylees. RCA is for people who:"
		  Text 30, 335, 275, 10, "- Have been in the United States eight months or less, and "
		  Text 30, 345, 275, 10, "- Have refugee or asylee status."

		  Text 305, 115, 225, 20, "Minnesota's Child Care Assistance Program makes quality child care affordable for families with low incomes, from the following programs:"
		  Text 310, 135, 225, 10, "MFIP Child Care"
		  Text 315, 145, 225, 30, "Families who receive assistance from the Diversionary Work Program or Minnesota Family Investment Program are eligible for child care if the parents are in work related activities."
		  Text 310, 175, 225, 10, "Transition Year Child Care"
		  Text 315, 185, 225, 30, "Available to families for up to 12 consecutive months after their Diversionary Work Program or Minnesota Family Investment Program case closes."
		  Text 310, 215, 225, 10, "Basic Sliding Fee Child Care"
		  Text 315, 225, 225, 10, "Available for other families with low incomes."
		  Text 310, 240, 225, 10, "Supplemental Nutrition Assistance Program (SNAP)"
		  Text 315, 250, 225, 30, "A federal program that helps Minnesotans with low income buy food. Benefits are available through EBT cards that can be used like money. Benefits are for:"
		  Text 320, 275, 225, 10, "- Single people"
		  Text 320, 285, 225, 10, "- Families with or without children"
		  Text 315, 295, 225, 20, "Your income, the size of your household, and your housing costs determines how much you can receive."


		  ButtonGroup ButtonPressed
		    PushButton 405, 340, 100, 13, "Open DHS 2920", open_program_info_doc
		  Text 10, 370, 210, 10, "Confirm you have reviewed Program Information:"
		  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! Program Information Discussed"+chr(9)+"No, I could not complete this", confirm_program_information_read
		EndDialog

		dialog Dialog1

		If ButtonPressed = open_program_info_doc Then
			err_msg = "LOOP"
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-2920-ENG"
		End If

		cancel_confirmation
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE


'DOMESTIC VIOLENCE INFORMATION - 3477
Do
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 550, 385, "Full Interview Questions"
		  ButtonGroup ButtonPressed
		    PushButton 465, 365, 80, 15, "Continue", continue_btn
		  Text 10, 5, 160, 10, "REVIEW the information listedd here to the client:"
		  GroupBox 10, 15, 530, 345, "Domestic Violence Information"
		  Text 15, 25, 505, 10, "If you are in danger from domestic violence or abuse and need help, call:"
		  Text 20, 35, 505, 10, "The National Domestic Violence Hotline at 800-799‑7233, (TTY: 800-787‑3224)"
		  Text 20, 45, 505, 10, "The Minnesota Coalition for Battered Women at 866-289‑6177"
		  Text 20, 55, 505, 10, "The Minnesota Day One Emergency Shelter and Crisis Hotline at 800-223‑1111"

		  Text 15, 65, 275, 10, "What is domestic violence?"
		  Text 20, 75, 275, 40, "Domestic violence or abuse is what someone says or does over and over again to make you feel afraid or to control you. People who are elderly, frail, have a disability, or who depend on others for assistance may not be able to protect themselves from domestic violence or abuse. Minnesota has a law to protect and assist people who are vulnerable to abuse or who are not able to care for themselves. Examples of violence or abuse include:"
		  Text 25, 115, 275, 10, "- Swearing or screaming at you"
		  Text 25, 125, 275, 10, "- Calling you names"
		  Text 25, 135, 275, 10, "- Taking money or property without permission"
		  Text 25, 145, 275, 10, "- Threatening to hurt you or others you care about"
		  Text 25, 155, 275, 10, "- Failing to provide care for you"
		  Text 25, 165, 275, 10, "- Not letting you leave your house"
		  Text 25, 175, 275, 10, "- Blaming you for everything that goes wrong"
		  Text 25, 185, 275, 10, "- Stalking you"
		  Text 25, 195, 275, 10, "- Being touched against your wishes or forced to have sex"
		  Text 25, 205, 275, 10, "- Choking, grabbing, hitting, pushing, pinching or kicking you."

		  Text 15, 215, 275, 20, "What services are available to victims of domestic violence or abuse?"

		  Text 20, 225, 275, 10, "Toll-free Hotlines have counselors who provide services, including:"
		  Text 25, 235, 275, 10, "- Crisis counseling"
		  Text 25, 245, 275, 10, "- Safety planning"
		  Text 25, 255, 275, 10, "- Assistance with finding shelter."
		  Text 20, 265, 275, 10, "Referrals to other organizations including:"
		  Text 25, 275, 275, 10, "- Legal services support groups"
		  Text 25, 285, 275, 10, "- Advocacy with the police."


		  Text 305, 65, 225, 10, "Safe At Home (SAH) Program"
		  Text 310, 75, 225, 60, "The Safe At Home (SAH) Program is a statewide address confidentiality program that assists survivors of domestic violence, sexual assault, stalking and others who fear for their safety by providing a substitute address for people who move or are about to move to a new location unknown to their aggressors. For information on this program, contact Safe At Home at 651-201‑1399 or 866-723‑3035."
		  Text 305, 135, 225, 10, "Vulnerable adults"
		  Text 310, 145, 225, 30, "Call the Senior LinkAge Line at 800-333-2433 to report concerns and to help a vulnerable adult get needed protection and assistance. Ask your worker for more resource information."
		  Text 305, 175, 275, 10, "What are domestic violence waivers?"
		  Text 310, 185, 275, 20, "If you are eligible for public assistance and you experience domestic violence, certain program requirements may not apply in your situation."
		  Text 310, 205, 275, 20, "If domestic violence or abuse makes it hard for you to follow program rules, talk to your county worker."


		  ButtonGroup ButtonPressed
		    PushButton 15, 340, 100, 13, "Open DHS 3477", open_DV_doc
		  Text 10, 370, 210, 10, "Confirm you have reviewed Domestic Violence Information:"
		  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! Domestic Violence Discussed"+chr(9)+"No, I could not complete this", confirm_DV_read
		EndDialog

		dialog Dialog1

		If ButtonPressed = open_DV_doc Then
			err_msg = "LOOP"
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG"
		End If

		cancel_confirmation
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE


'DO YOU HAVE A DISABILITY - 4133
Do
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 550, 385, "Full Interview Questions"
		  ButtonGroup ButtonPressed
		    PushButton 465, 365, 80, 15, "Continue", continue_btn
		  Text 10, 5, 160, 10, "REVIEW the information listedd here to the client:"
		  GroupBox 10, 15, 530, 345, "Do you have a Disability"
		  Text 15, 25, 505, 10, "Please tell us if you have a disability so we can help you access human services programs and benefits."

		  Text 15, 35, 275, 10, "What medical conditions may be disabilities?"
		  Text 20, 45, 275, 20, "A disability is a physical, sensory, or mental impairment that materially limits a major life activity. Types of disabilities may include:"
		  Text 25, 65, 275, 10, "- Diseases like diabetes, epilepsy or cancer"
		  Text 25, 75, 275, 10, "- Learning disorders like dyslexia"
		  Text 25, 85, 275, 10, "- Developmental delays"
		  Text 25, 95, 275, 10, "- Clinical depression"
		  Text 25, 105, 275, 10, "- Hearing loss or low vision"
		  Text 25, 115, 275, 10, "- Movement restrictions like trouble with walking, reaching or grasping"
		  Text 25, 125, 275, 10, "- History of alcohol or drug addiction"
		  Text 35, 135, 275, 10, "(current illegal drug use is not a disability)"
		  Text 20, 145, 275, 30, "If you are asking for or are getting benefits through either a county human services agency or the Minnesota Department of Human Services, that office will let you know if you have a disability using information from you and your doctor."

		  Text 305, 35, 225, 10, "What help is available?"
		  Text 310, 45, 225, 20, "If you have a disability, your county or the state human services agency can help you by:"
		  Text 315, 65, 225, 20, "- Calling you or meeting with you in another place if you are not able to come into the office"
		  Text 315, 85, 225, 10, "- Using a sign language interpreter"
		  Text 315, 95, 225, 20, "- Giving you letters and forms in other formats like computer files, audio recordings, large print or Braille"
		  Text 315, 115, 225, 10, "- Telling you the meaning of the information we give you"
		  Text 315, 125, 225, 10, "- Helping you fill out forms"
		  Text 315, 135, 225, 10, "- Helping you make a plan so you can work even with your disability"
		  Text 315, 145, 225, 10, "- Sending you to other services that may help you"
		  Text 315, 155, 225, 20, "- Helping you to appeal agency decisions about you if you disagree with them"
		  Text 310, 175, 225, 30, "You will not have to pay extra for help. If you want help, ask your agency as soon as possible. An agency may not be able to accommodate requests made within 48 hours of need."

		  Text 15, 205, 505, 10, "How does the law protect people with disabilities?"
		  Text 20, 215, 505, 40, "The Americans with Disabilities Act (ADA) and the ADA Amendments Act are federal laws, and the Minnesota Human Rights Act is a state law. Each gives individuals with disabilities the same legal rights and protections as people without disabilities, including access to public assistance benefits. You will not be denied benefits because you have a disability. Your benefits will not be stopped because of your disability. If your disability makes getting benefits hard for you, your county human services agency will help you access all of the programs that are available to you."

		  ButtonGroup ButtonPressed
		    PushButton 15, 340, 100, 13, "Open DHS 4133", open_disa_doc
		  Text 10, 370, 210, 10, "Confirm you have reviewed Domestic Violence Information:"
		  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! Disability Information Discussed"+chr(9)+"No, I could not complete this", confirm_disa_read
		EndDialog

		dialog Dialog1

		If ButtonPressed = open_disa_doc Then
			err_msg = "LOOP"
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-4133-ENG"
		End If

		cancel_confirmation
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE




'MFIP CASES
	'Reporting Responsibilities for MFIP Households (DHS-2647) (PDF).
	'Notice of Requirement to Attend MFIP Overview (DHS-2929) (PDF). See 0028.09 (ES Overview/SNAP E&T Orientation).
	'Family Violence Referral (DHS-3323) (PDF) and
Do
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 550, 385, "Full Interview Questions"

		  Text 10, 5, 160, 10, "REVIEW the information listedd here to the client:"
		  GroupBox 10, 15, 530, 345, "MFIP Cases"

		  GroupBox 10, 25, 530, 95, "Reporting Responsibilities for MFIP Households (DHS-2647)"
		  Text 15, 40, 505, 10, "Changes you must report: Anything that could impact eligibility. Particularly: Income, Assets, Household Comp"
		  Text 15, 50, 505, 10, "When do Changes need to be Reported: On the monthly Household Report Form, if you do not have one, within 10 days of the change"
		  Text 15, 60, 505, 10, "How to report Changes: On any Report Form or call the county."
		  GroupBox 10, 125, 530, 95, "Notice of Requirement to Attend MFIP Overview (DHS-2929)"
		  Text 15, 140, 505, 10, "All MFIP caregivers are required to attend an MFIP overview and participate in Employment Services."
		  Text 15, 150, 505, 10, "If you do not go to your scheduled overview meeting without good reason, your MFIP grant may be reduced until you go to the meeting."
		  Text 15, 160, 505, 10, "Call the contact person above if you: - Need child care or help getting to the meeting - Have problems attending the meeting."
		  GroupBox 10, 225, 530, 95, "Family Violence Referral (DHS-3323)"
		  Text 15, 240, 505, 10, "If you, or someone in your home is a victim of domestic abuse the county can help you."
		  Text 15, 250, 505, 10, "You can also call the National Domestic Violence Hot Line at (800) 799-7233 or Legal Aid at (888) 354-5522."
		  Text 15, 260, 505, 20, "Some of the Minnesota Family Investment Program (MFIP) rules do not apply to domestic abuse victims. You must tell us about the abuse and have a special employment plan that includes activities to help keep your family safe. Please talk to your worker or an advocate if you want to know about this."
		  ButtonGroup ButtonPressed
		    PushButton 465, 365, 80, 15, "Continue", continue_btn
		    PushButton 430, 22, 100, 13, "Open DHS 2647", open_cs_2647_doc
			PushButton 430, 122, 100, 13, "Open DHS 2929", open_cs_2929_doc
			PushButton 430, 222, 100, 13, "Open DHS 3323", open_cs_3323_doc
		  Text 10, 370, 210, 10, "Confirm you have reviewed MFIP Specific Information:"
		  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! MFIP Forms Discussed"+chr(9)+"No, I could not complete this", confirm_mfip_forms_read
		EndDialog

		dialog Dialog1

		If ButtonPressed = open_cs_2647_doc OR ButtonPressed = open_cs_2929_doc OR ButtonPressed = open_cs_3323_doc Then
			err_msg = "LOOP"
			If ButtonPressed = open_cs_2647_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-2647-ENG"
			If ButtonPressed = open_cs_2929_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-2929-ENG"
			If ButtonPressed = open_cs_3323_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-3323-ENG"
		End If

		cancel_confirmation
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE


'In cases where there is at least 1 non-custodial parent:
	'Understanding Child Support - A Handbook for Parents (DHS-3393) (PDF).
	'Referral to Support and Collections (DHS-3163B) (PDF). (This is in addition to the Combined Application Form or ApplyMN application, for EACH non-custodial parent). See 0012.21.03 (Support From Non-Custodial Parents).
	'Cooperation with Child Support Enforcement (DHS-2338) (PDF). See 0012.21.06 (Child Support Good Cause Exemptions).
'If a non-parental caregiver applies,
	'MFIP Child Only Assistance (DHS-5561) (PDF).
Do
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 550, 385, "Full Interview Questions"
		  Text 10, 5, 160, 10, "REVIEW the information listedd here to the client:"
		  GroupBox 10, 15, 530, 345, "MFIP Case with at least 1 ABPS - Child Support Information"

		  GroupBox 10, 25, 530, 95, "Understanding Child Support - A Handbook for Parents (DHS-3393)"
		  Text 15, 40, 505, 20, "Every child needs financial and emotional support. Every child has the right to this support from both parents. Devoted parents can be loving and supportive forces in a child's life. Even when parents do not live together, they need to work together to support their child."
		  Text 15, 60, 505, 20, "Minnesoda Child Support and Hennepin County Child Support provide support and guidance. The Handbook 'Understanding Child Support' provides information about the details of these programs."
		  GroupBox 10, 125, 530, 95, "Referral to Support and Collections (DHS-3163B)"
		  Text 15, 140, 505, 10, "Purpose of form: The child support agency will use the information you give to help collect support."
		  Text 15, 150, 505, 20, "How to complete this form: Fill in each blank. If there are boxes, check the box or boxes that fit your situation. Complete a separate form for each parent or alleged parent other than yourself."
		  Text 15, 170, 505, 20, "Please read the booklet 'Understanding Child Support: A Handbook for Parents' (DHS-3393) before signing. The booklet explains information about the child support services you may be receiving."
		  GroupBox 10, 225, 530, 95, "Cooperation with Child Support Enforcement (DHS-2338)"
		  Text 15, 240, 505, 10, "This notice explains your rights and responsibilities for cooperating with the MN Department of Human Services, Child Support Division."
		  Text 15, 250, 505, 10, "Cooperation with the child support agency includes answering questions, filling out forms, and appearing at appointments and/or court hearings."
		  Text 15, 260, 505, 40, "This notice also explains how you make a 'good cause claim' that gives you the right not to cooperate if your claim is granted. If you choose to claim good cause and your county child support agency is currently collecting your child support payments, the county will immediately stop collecting those payments for the child(ren) you name on the attached form. The county will stop providing all child support services until it makes a decision on your good cause claim. If you are granted a good cause exemption, the child support agency will close your case."
		  GroupBox 10, 325, 530, 40, "If Non-Custodial Caregiver - MFIP Child Only Assistance (DHS-5561)"
		  Text 15, 335, 505, 30, "The Minnesota Department of Human Services has assistance programs available to help children who are cared for and supported by their relatives. This brochure answers some frequently asked questions relatives may have about the Minnesota Family Investment Program (MFIP)"
		  ButtonGroup ButtonPressed
		    PushButton 465, 365, 80, 15, "Continue", continue_btn
		    PushButton 430, 22, 100, 13, "Open DHS 3393", open_cs_3393_doc
			PushButton 430, 122, 100, 13, "Open DHS 3163B", open_cs_3163B_doc
			PushButton 430, 222, 100, 13, "Open DHS 2338", open_cs_2338_doc
			PushButton 430, 322, 100, 13, "Open DHS 5561", open_cs_5561_doc

		  Text 10, 370, 210, 10, "Confirm you have reviewed MFIP Child Support Information:"
		  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! MFIP Child Support Discussed"+chr(9)+"No, I could not complete this", confirm_mfip_cs_read
		EndDialog

		dialog Dialog1

		If ButtonPressed = open_cs_3393_doc OR ButtonPressed = open_cs_3163B_doc OR ButtonPressed = open_cs_2338_doc OR ButtonPressed = open_cs_5561_doc Then
			err_msg = "LOOP"
			If ButtonPressed = open_cs_3393_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-3393-ENG"
			If ButtonPressed = open_cs_3163B_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-3163B-ENG"
			If ButtonPressed = open_cs_2338_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-2338-ENG"
			If ButtonPressed = open_cs_5561_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-5561-ENG"
		End If

		cancel_confirmation
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

'If there is a custodial parent under 20, the
	'Notice of Requirement to Attend School (DHS-2961) (PDF) and
	'Graduate to Independence - MFIP Teen Parent Informational Brochure (DHS-2887) (PDF).
'If there is a custodial parent under age 18, the
	'MFIP for Minor Caregivers (DHS-3238) (PDF) brochure.
Do
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 550, 385, "Full Interview Questions"
		  Text 10, 5, 160, 10, "REVIEW the information listedd here to the client:"
		  GroupBox 10, 15, 530, 345, "MFIP Case Minor Caregiver Cases"

		  GroupBox 10, 25, 530, 95, "Notice of Requirement to Attend School (DHS-2961)"
		  Text 15, 40, 505, 10, "This form tells you that, unless you are exempt, you must attend school and what will happen if you do not go to school."
		  Text 15, 50, 505, 20, "The first step is for us to complete an assessment with you. We will review your educational progress, needs, literacy level, family circumstances, skills, and work experience. We will see if you need child care or other services so you can go to school."
		  Text 15, 70, 505, 10, "If you do not cooperate or do not attend school, without good cause, we will send you a notice. This notice will tell you that your MFIP grant may be reduced. "
		  GroupBox 10, 125, 530, 95, "Graduate to Independence - MFIP Teen Parent Informational Brochure (DHS-2887)"
		  Text 15, 140, 505, 20, "If you are a teen parent under the age of 20, and do not have a high school diploma or an equivalent, you are expected to attend an approved educational program to qualify for the Minnesota Family Investment Program."
		  Text 15, 160, 505, 20, "Earning your diploma is the first step in getting ready for a job. County human services staff will help you with counseling, child care, and transportation so you can go to school. They will also help you find a school program that is best for you."
		  Text 15, 180, 505, 10, "If you fail to attend school, without good cause, your human services worker will reduce your grant by 10 percent or more of your standard of need."
		  GroupBox 10, 225, 530, 95, "MFIP for Minor Caregivers (DHS-3238)"
		  Text 15, 240, 505, 20, "You are a minor caregiver if: "
		  Text 25, 250, 505, 20, "- You are younger than 18 - You have never been married - You are not emancipated and - You are the parent of a child(ren) living in the same household."
		  Text 15, 260, 505, 10, "If you are a minor caregiver, to receive benefits and services, you must be living: "
		  Text 25, 270, 505, 20, "- With a parent or with an adult relative caregiver or with a legal guardian or - In an agency-approved living arrangement."
		  Text 15, 280, 505, 10, "A social worker must approve any exception(s) to your living arrangement."
		  ButtonGroup ButtonPressed
		  	PushButton 465, 365, 80, 15, "Continue", continue_btn
		    PushButton 430, 22, 100, 13, "Open DHS 2961", open_cs_2961_doc
			PushButton 430, 122, 100, 13, "Open DHS 2887", open_cs_2887_doc
			PushButton 430, 222, 100, 13, "Open DHS 3238", open_cs_3238_doc

		  Text 10, 370, 210, 10, "Confirm you have reviewed MFIP Minor Caregiver Information:"
		  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! MFIP Minor Caregiver Discussed"+chr(9)+"No, I could not complete this", confirm_minor_mfip_read
		EndDialog

		dialog Dialog1

		If ButtonPressed = open_cs_2961_doc OR ButtonPressed = open_cs_2887_doc OR ButtonPressed = open_cs_3238_doc Then
			err_msg = "LOOP"
			If ButtonPressed = open_cs_2961_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://edocs.dhs.state.mn.us/lfserver/Legacy/DHS-2961-ENG"
			If ButtonPressed = open_cs_2887_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-2887-ENG"
			If ButtonPressed = open_cs_3238_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-3238-ENG"
		End If

		cancel_confirmation
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE




'SNAP CASES'
	'Supplemental Nutrition Assistance Program reporting responsibilities (DHS-2625).
	'Facts on Voluntarily Quitting Your Job If You Are on the Supplemental Nutrition Assistance Program (SNAP) (DHS-2707).
	'Work Registration Notice (DHS-7635).
Do
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 550, 385, "Full Interview Questions"
		  Text 10, 5, 160, 10, "REVIEW the information listedd here to the client:"
		  GroupBox 10, 15, 530, 345, "SNAP Case"

		  GroupBox 10, 25, 530, 95, "Supplemental Nutrition Assistance Program reporting responsibilities (DHS-2625)"
		  Text 15, 40, 505, 10, "There are the three reporting types used by SNAP: Six-month reporting, Change reporting, Monthly reporting"
		  Text 15, 50, 505, 10, "You are _____________ Reporting. This means you must report:"
		  Text 15, 60, 505, 10, "Complete a renewal every six months. Report changes in income over 130% FPG ($XXXX) by the 10th of the month followwing the month of change."
		  GroupBox 10, 125, 530, 95, "Facts on Voluntarily Quitting Your Job If You Are on SNAP (DHS-2707)"
		  Text 15, 140, 505, 10, "If you or someone else in your household has a job and quits without a good reason, your household might not get SNAP benefits."
		  Text 15, 150, 505, 20, "The penalty does not apply if the person who quit a job: "
		  Text 25, 160, 505, 20, "- Was fired, or forced to leave the job, or had hours cut back by the employer - Was self-employed - Left a job that was less than 30 hours per week"
		  Text 15, 170, 505, 10, "The penalty also does not apply if you can prove the person had 'good reason' to quit the job. The form has some examples of 'good reasons'."
		  GroupBox 10, 225, 530, 95, "Work Registration Notice (DHS-7635)"
		  Text 15, 240, 505, 10, "In order to be eligible for benefits you must cooperate in any efforts regarding work registration. "
		  Text 15, 250, 505, 10, "If you do not follow any of the work requirements listed above your benefit smay end."
		  ButtonGroup ButtonPressed
		    PushButton 465, 365, 80, 15, "Continue", continue_btn
		    PushButton 430, 22, 100, 13, "Open DHS 2625", open_cs_2625_doc
			PushButton 430, 122, 100, 13, "Open DHS 2707", open_cs_2707_doc
			PushButton 430, 222, 100, 13, "Open DHS 7635", open_cs_7635_doc

		  Text 10, 370, 210, 10, "Confirm you have reviewed SNAP Specific Information:"
		  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! SNAP Forms Discussed"+chr(9)+"No, I could not complete this", confirm_snap_forms_read
		EndDialog

		dialog Dialog1

		If ButtonPressed = open_cs_2625_doc OR ButtonPressed = open_cs_2707_doc OR ButtonPressed = open_cs_7635_doc Then
			err_msg = "LOOP"
			If ButtonPressed = open_cs_2625_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-2625-ENG"
			If ButtonPressed = open_cs_2707_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-2707-ENG"
			If ButtonPressed = open_cs_7635_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-7635-ENG"
		End If

		cancel_confirmation
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

'Employment Services Registration.

'REPORTING

'Additional Important Information.

'Penalty Warnings.



Dialog1 = ""
BeginDialog Dialog1, 0, 0, 206, 265, "Resources MEMO"
  ButtonGroup ButtonPressed
    PushButton 150, 25, 50, 10, "Check All", check_all_button
  CheckBox 10, 40, 145, 10, "Client Email Submission/Virtual Dropbox", client_virtual_dropox_checkbox
  CheckBox 10, 55, 140, 10, "Community Action Partnership - CAP", cap_checkbox
  CheckBox 10, 70, 115, 10, "DHS MMIS Recipient HelpDesk", MMIS_helpdesk_checkbox
  CheckBox 10, 85, 180, 10, "DHS MNSure Helpdesk   * NOT FOR MA CLIENTS", MNSURE_helpdesk_checkbox
  CheckBox 10, 100, 145, 10, "Disability Hub (Disability Linkage Line)", disability_hub_checkbox
  CheckBox 10, 115, 125, 10, "Emergency Mental Health Services", emer_mental_health_checkbox
  CheckBox 10, 130, 175, 10, "Emergency Food Shelf Network (The Food Group)", emer_food_network_checkbox
  CheckBox 10, 145, 50, 10, "Front Door", front_door_checkbox
  CheckBox 10, 160, 75, 10, "Senior Linkage Line", sr_linkage_line_checkbox
  CheckBox 10, 175, 130, 10, "United Way First Call for Help (211)", united_way_checkbox
  CheckBox 10, 190, 60, 10, "Xcel Energy", xcel_checkbox
  Text 5, 5, 195, 20, "Does the client need any additional resources or supports. Check any that the client may need."
  ButtonGroup ButtonPressed
    PushButton 150, 245, 50, 15, "Continue", continue_btn
  Text 10, 210, 185, 35, "When you press continue, the script will display these resources for you to give them verbally to the client. It will then send a MEMO or create a Word Doc to provide to the client in writing."
EndDialog
'This Do...loop shows the appointment letter dialog, and contains logic to require most fields.

DO
	Do
		err_msg = ""
		Dialog Dialog1
		' If ButtonPressed = cancel then stopscript
        ' If cap_checkbox = unchecked AND emer_mental_health_checkbox = unchecked AND MMIS_helpdesk_checkbox = unchecked AND MNSURE_helpdesk_checkbox = unchecked AND disability_hub_checkbox = unchecked AND emer_food_network_checkbox = unchecked AND front_door_checkbox = unchecked AND sr_linkage_line_checkbox = unchecked AND united_way_checkbox = unchecked AND xcel_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "You must select at least one resource."
		' If isnumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & "You must fill in a valid case number." & vbNewLine
		' If worker_signature = "" then err_msg = err_msg & "You must sign your case note." & vbNewLine
        ' If ButtonPressed = check_all_button Then
        '     err_msg = "LOOP" & err_msg
		'
		' 	client_virtual_dropox_checkbox = checked
        '     cap_checkbox = checked
        '     MMIS_helpdesk_checkbox = checked
        '     MNSURE_helpdesk_checkbox = checked
        '     disability_hub_checkbox = checked
        '     emer_food_network_checkbox = checked
        '     emer_mental_health_checkbox = checked
        '     front_door_checkbox = checked
        '     sr_linkage_line_checkbox = checked
        '     united_way_checkbox = checked
        '     xcel_checkbox = checked
        ' End If
		' IF err_msg <> "" AND left(err_msg, 4) <> "LOOP" THEN msgbox err_msg
	Loop until err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false



Do
	Do
		err_msg = ""

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 550, 385, "Full Interview Questions"
		  Text 10, 10, 500, 10, "The Interview Information has been completed. Review the information and next steps with the client."
		  GroupBox 10, 20, 530, 340, "CASE INTERVIEW WRAP UP"

		  Text 15, 30, 505, 10, "Programs being Requested/Renewed:"
		  Text 20, 40, 505, 10, "SNAP"
		  Text 20, 50, 505, 10, "Cash - MFIP"
		  Text 20, 60, 505, 10, "Housing Support - GRH"
		  Text 15, 75, 505, 10, "Next Steps:"
		  Text 20, 85, 505, 10, "We need verifications before we can make a determination on your case. Are you clear on what those are? You will also receive a notice in the mail."
		  Text 20, 95, 505, 10, "If you need an EBT Card - call or go in."
		  Text 20, 105, 505, 10, "I will be processing your case. "
		  Text 25, 115, 505, 10, "APPLICATION - the benefits are typically available the day after appproval. "
		  Text 25, 125, 505, 10, "RECERT - the benefits should be available on your regular day."
		  Text 20, 135, 505, 10, "Watch your mail for approval notices to see the benefit amount."
		  Text 15, 150, 505, 10, "Your address and phone number are our best way to contact you."
		  Text 20, 160, 505, 10, "It is vital that you let us know if you address or phone number has changed"
		  Text 20, 170, 505, 10, "You may miss important requests or notices if we have an old address."
		  Text 20, 180, 505, 10, "Our mail does not forward to address changes, so we need to know the correct address for you"
		  Text 15, 195, 505, 10, "Please be sure to follow program rules and requirements"
		  Text 20, 205, 505, 10, "Failure to report changes and information timely can have negative impacts:"
		  Text 25, 215, 505, 10, "End of benefits"
		  Text 25, 225, 505, 10, "Overpayments"
		  Text 25, 235, 505, 10, "Future ineligibility"
		  Text 20, 245, 505, 10, "We receive information from other sources about you and may impact your eligibility and benefit level."
		  Text 20, 255, 505, 10, "If you are unsure of program rules and requirements, the forms we reviewed earlier can always be resent, or you can call us with questions."
		  Text 15, 270, 505, 10, "Contact to Hennepin County"
		  Text 20, 280, 505, 10, "By Phone - 612-596-1300. The phone lines are open Monday - Friday 8:00 - 4:30"
		  Text 20, 290, 505, 10, "In person - at one of six regional hubs"
		  Text 20, 300, 505, 10, "Online - InfoKeep"

		  ButtonGroup ButtonPressed
		    PushButton 465, 365, 80, 15, "Continue", continue_btn
		    ' PushButton 430, 22, 100, 13, "Open DHS 2625", open_cs_2625_doc
			' PushButton 430, 122, 100, 13, "Open DHS 2707", open_cs_2707_doc
			' PushButton 430, 222, 100, 13, "Open DHS 7635", open_cs_7635_doc

		  Text 10, 370, 210, 10, "Confirm you have reviewed Hennepin County Information Information:"
		  DropListBox 220, 365, 175, 45, "Enter confirmation"+chr(9)+"YES! Recap Discussed"+chr(9)+"No, I could not complete this", confirm_recap_read
		EndDialog

		dialog Dialog1

		' If ButtonPressed = open_cs_2625_doc OR ButtonPressed = open_cs_2707_doc OR ButtonPressed = open_cs_7635_doc Then
		' 	err_msg = "LOOP"
		' 	If ButtonPressed = open_cs_2625_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-2625-ENG"
		' 	If ButtonPressed = open_cs_2707_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-2707-ENG"
		' 	If ButtonPressed = open_cs_7635_doc Then run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe http://edocs.dhs.state.mn.us/lfserver/Public/DHS-7635-ENG"
		' End If

		cancel_confirmation
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

interview_time = timer - start_time
interview_time = interview_time/60
interview_time = Round(interview_time, 2)
complete_interview_msg = MsgBox("This interview is now completed and has taken " & interview_time & " minutes." & vbCr & vbCr & "The script will now create your interview notes in a PDF and enter CASE:NOTE(s) as needed.", vbInformation, "Interview Completed")

' script_end_procedure("At this point the script will create a PDF with all of the interview notes to save to ECF, enter a comprehensive CASE:NOTE, and update PROG or REVW with the interview date. Future enhancements will add more actions functionality.")
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
objSelection.TypeText "Interview Date: " & interview_date & vbCR
objSelection.TypeText "DATE OF APPLICATION: " & CAF_datestamp & vbCR
objSelection.TypeText "Completed by: " & worker_name & vbCR
objSelection.TypeText "Interview completed with: " & who_are_we_completing_the_interview_with & vbCR
length_of_interview = (timer-start_time)/60
objSelection.TypeText "Interview length: " & length_of_interview & " minutes" & vbCR
objSelection.TypeText "Case Status at the time of interview: " & vbCR
If case_active = TRUE Then
	objSelection.TypeText "   Case is ACTIVE" & vbCR
ElseIf case_pending = TRUE Then
	objSelection.TypeText "   Case is PENDING" & vbCR
Else
	objSelection.TypeText "   Case is INACTIVE" & vbCR
End If

Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 7, 3					'This sets the rows and columns needed row then column
'This table starts with 1 column - other columns are added after we split some of the cells
set objProgStatusTable = objDoc.Tables(1)		'Creates the table with the specific index'

objProgStatusTable.AutoFormat(16)							'This adds the borders to the table and formats it
objProgStatusTable.Columns(1).Width = 150					'This sets the width of the table.
objProgStatusTable.Columns(2).Width = 200					'This sets the width of the table.
objProgStatusTable.Columns(3).Width = 150					'This sets the width of the table.

' objProgStatusTable.Cell(row, col).Range.Text =

objProgStatusTable.Cell(1, 1).Range.Text = "Program"
objProgStatusTable.Cell(1, 2).Range.Text = "Status"
objProgStatusTable.Cell(1, 3).Range.Text = "Detail"

objProgStatusTable.Cell(2, 1).Range.Text = "SNAP"
objProgStatusTable.Cell(2, 2).Range.Text = snap_case
' objProgStatusTable.Cell(2, 3).Range.Text =
cash_col = 3
' If
If mfip_case = True Then
	If cash_col = 3 Then objProgStatusTable.Cell(3, 1).Range.Text = "CASH 1"
	If cash_col = 4 Then objProgStatusTable.Cell(4, 1).Range.Text = "CASH 2"
	objProgStatusTable.Cell(cash_col, 2).Range.Text = mfip_case
	objProgStatusTable.Cell(cash_col, 3).Range.Text = "MFIP"
	cash_col = cash_col + 1
End If
If dwp_case = True Then
	If cash_col = 3 Then objProgStatusTable.Cell(3, 1).Range.Text = "CASH 1"
	If cash_col = 4 Then objProgStatusTable.Cell(4, 1).Range.Text = "CASH 2"
	objProgStatusTable.Cell(cash_col, 2).Range.Text = dwp_case
	objProgStatusTable.Cell(cash_col, 3).Range.Text = "DWP"
	cash_col = cash_col + 1
End If
If ga_case = True Then
	If cash_col = 3 Then objProgStatusTable.Cell(3, 1).Range.Text = "CASH 1"
	If cash_col = 4 Then objProgStatusTable.Cell(4, 1).Range.Text = "CASH 2"
	objProgStatusTable.Cell(cash_col, 2).Range.Text = ga_case
	objProgStatusTable.Cell(cash_col, 3).Range.Text = "GA"
	cash_col = cash_col + 1
End If
If msa_case = True Then
	If cash_col = 3 Then objProgStatusTable.Cell(3, 1).Range.Text = "CASH 1"
	If cash_col = 4 Then objProgStatusTable.Cell(4, 1).Range.Text = "CASH 2"
	objProgStatusTable.Cell(cash_col, 2).Range.Text = msa_case
	objProgStatusTable.Cell(cash_col, 3).Range.Text = "MSA"
	cash_col = cash_col + 1
End If
If unknown_cash_pending = True Then
	If cash_col = 3 Then objProgStatusTable.Cell(3, 1).Range.Text = "CASH 1"
	If cash_col = 4 Then objProgStatusTable.Cell(4, 1).Range.Text = "CASH 2"
	objProgStatusTable.Cell(cash_col, 2).Range.Text = unknown_cash_pending
	objProgStatusTable.Cell(cash_col, 3).Range.Text = "CASH"
	cash_col = cash_col + 1
End If
If cash_col = 3 Then
	objProgStatusTable.Cell(3, 1).Range.Text = "CASH 1"
	objProgStatusTable.Cell(cash_col, 2).Range.Text = "NONE"
	objProgStatusTable.Cell(cash_col, 3).Range.Text = ""
	cash_col = cash_col + 1
End If
If cash_col = 4 Then
	objProgStatusTable.Cell(4, 1).Range.Text = "CASH 2"
	objProgStatusTable.Cell(cash_col, 2).Range.Text = "NONE"
	objProgStatusTable.Cell(cash_col, 3).Range.Text = ""
	cash_col = cash_col + 1
End If

objProgStatusTable.Cell(5, 1).Range.Text = "GRH"
objProgStatusTable.Cell(5, 2).Range.Text = grh_case

objProgStatusTable.Cell(6, 1).Range.Text = "MA"
objProgStatusTable.Cell(6, 2).Range.Text = ma_case

objProgStatusTable.Cell(7, 1).Range.Text = "MSA"
objProgStatusTable.Cell(7, 2).Range.Text = msp_case

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing

'Program CAF Information
caf_progs = ""
for the_memb = 0 to UBOUND(HH_MEMB_ARRAY, 2)
	If HH_MEMB_ARRAY(snap_req_checkbox, the_memb) = checked AND InStr(caf_progs, "SNAP") = 0 Then caf_progs = caf_progs & ", SNAP"
	If HH_MEMB_ARRAY(cash_req_checkbox, the_memb) = checked AND InStr(caf_progs, "Cash") = 0 Then caf_progs = caf_progs & ", Cash"
	If HH_MEMB_ARRAY(emer_req_checkbox, the_memb) = checked AND InStr(caf_progs, "EMER") = 0 Then caf_progs = caf_progs & ", EMER"
Next
If left(caf_progs, 2) = ", " Then caf_progs = right(caf_progs, len(caf_progs)-2)
objSelection.TypeText "PROGRAMS REQUESTED ON CAF: " & caf_progs & vbCr
objSelection.Font.Size = "11"


'Ennumeration for SetHeight and SetWidth
'wdAdjustFirstColumn	2	Adjusts the left edge of the first column only, preserving the positions of the other columns and the right edge of the table.
	' wdAdjustNone			0	Adjusts the left edge of row or rows, preserving the width of all columns by shifting them to the left or right. This is the default value.
	' wdAdjustProportional	1	Adjusts the left edge of the first column, preserving the position of the right edge of the table by proportionally adjusting the widths of all the cells in the specified row or rows.
	' wdAdjustSameWidth		3	Adjusts the left edge of the first column, preserving the position of the right edge of the table by setting the widths of all the cells in the specified row or rows to the same value.


objSelection.TypeText "PERSON 1 Information - Confirmed in the Interview"
Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 16, 1					'This sets the rows and columns needed row then column
'This table starts with 1 column - other columns are added after we split some of the cells
set objPers1Table = objDoc.Tables(2)		'Creates the table with the specific index'
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
objPers1Table.Cell(2, 1).Range.Text = HH_MEMB_ARRAY(last_name_const, 0)
objPers1Table.Cell(2, 2).Range.Text = HH_MEMB_ARRAY(first_name_const, 0)
objPers1Table.Cell(2, 3).Range.Text = HH_MEMB_ARRAY(mid_initial, 0)
objPers1Table.Cell(2, 4).Range.Text = HH_MEMB_ARRAY(other_names, 0)

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

objPers1Table.Cell(4, 1).Range.Text = HH_MEMB_ARRAY(ssn, 0)
objPers1Table.Cell(4, 2).Range.Text = HH_MEMB_ARRAY(date_of_birth, 0)
objPers1Table.Cell(4, 3).Range.Text = HH_MEMB_ARRAY(gender, 0)
objPers1Table.Cell(4, 4).Range.Text = HH_MEMB_ARRAY(marital_status, 0)

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
objPers1Table.Cell(5, 1).Range.Text = "RESIDENCE ADDRESS - Confirmed in the Interview"
' objPers1Table.Cell(5, 2).Range.Text = "APT. NUMBER"
objPers1Table.Cell(5, 2).Range.Text = "CITY"
objPers1Table.Cell(5, 3).Range.Text = "STATE"
objPers1Table.Cell(5, 4).Range.Text = "ZIP CODE"

If homeless_yn = "Yes" Then
	objPers1Table.Cell(6, 1).Range.Text = resi_addr_street_full & " - HOMELESS - "
Else
	objPers1Table.Cell(6, 1).Range.Text = resi_addr_street_full
End If
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

objPers1Table.Cell(12, 1).Range.Text = HH_MEMB_ARRAY(interpreter, 0)
objPers1Table.Cell(12, 2).Range.Text = HH_MEMB_ARRAY(spoken_lang, 0)
objPers1Table.Cell(12, 3).Range.Text = HH_MEMB_ARRAY(written_lang, 0)

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

objPers1Table.Cell(14, 1).Range.Text = HH_MEMB_ARRAY(last_grade_completed, 0)
objPers1Table.Cell(14, 2).Range.Text = "Date: " & HH_MEMB_ARRAY(mn_entry_date, 0) & "   From: " & HH_MEMB_ARRAY(former_state0)
objPers1Table.Cell(14, 3).Range.Text = HH_MEMB_ARRAY(citizen, 0)

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
If HH_MEMB_ARRAY(none_req_checkbox, 0) = checked then progs_applying_for = "NONE"
If HH_MEMB_ARRAY(snap_req_checkbox, 0) = checked then progs_applying_for = progs_applying_for & ", SNAP"
If HH_MEMB_ARRAY(cash_req_checkbox, 0) = checked then progs_applying_for = progs_applying_for & ", Cash"
If HH_MEMB_ARRAY(emer_req_checkbox, 0) = checked then progs_applying_for = progs_applying_for & ", Emergency Assistance"
If left(progs_applying_for, 2) = ", " Then progs_applying_for = right(progs_applying_for, len(progs_applying_for) - 2)

'defining a string of the races that were selected from checkboxes in the dialog.
If HH_MEMB_ARRAY(race_a_checkbox, 0) = checked then race_to_enter = race_to_enter & ", Asian"
If HH_MEMB_ARRAY(race_b_checkbox, 0) = checked then race_to_enter = race_to_enter & ", Black"
If HH_MEMB_ARRAY(race_n_checkbox, 0) = checked then race_to_enter = race_to_enter & ", American Indian or Alaska Native"
If HH_MEMB_ARRAY(race_p_checkbox, 0) = checked then race_to_enter = race_to_enter & ", Pacific Islander and Native Hawaiian"
If HH_MEMB_ARRAY(race_w_checkbox, 0) = checked then race_to_enter = race_to_enter & ", White"
If left(race_to_enter, 2) = ", " Then race_to_enter = right(race_to_enter, len(race_to_enter) - 2)

objPers1Table.Cell(16, 1).Range.Text = progs_applying_for
objPers1Table.Cell(16, 2).Range.Text = HH_MEMB_ARRAY(ethnicity_yn, 0)
objPers1Table.Cell(16, 3).Range.Text = race_to_enter

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
objSelection.TypeParagraph()						'adds a line between the table and the next information

objSelection.TypeText "LIVING SITUATION: " & living_situation & vbCR
objSelection.TypeText "INTERVIEW NOTES: " & HH_MEMB_ARRAY(client_notes, 0) & vbCR

' objSelection.Font.Bold = TRUE
objSelection.TypeText "CAF 1 - EXPEDITED QUESTIONS from the CAF"
Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 8, 2					'This sets the rows and columns needed row then column'
set objEXPTable = objDoc.Tables(3)		'Creates the table with the specific index'

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
objSelection.TypeParagraph()						'adds a line between the table and the next information
objSelection.Font.Bold = TRUE
objSelection.TypeText "EXPEDITED Interview Answers:" & vbCr
objSelection.Font.Bold = FALSE
If appears_expedited = True Then
	objSelection.TypeText "Based on income information this case APPEARS ELIGIBLE FOR EXPEDITED SNAP." & vbCr
Else
	objSelection.TypeText "This case does not appear eligible for expedited SNAP based on the income information." & vbCr
End If
total_expenses = 0
utilities_cost = 0
If intv_exp_pay_heat_checkbox = checked OR intv_exp_pay_ac_checkbox = checked Then
	utilities_cost = 496
Else
	If intv_exp_pay_electricity_checkbox = checked Then utilities_cost = utilities_cost + 154
	If intv_exp_pay_phone_checkbox = checked Then utilities_cost = utilities_cost + 56
End If
total_expenses = intv_app_month_housing_expense + utilities_cost

objSelection.TypeText chr(9) & "Income in the month of application: " & intv_app_month_income & vbCr
objSelection.TypeText chr(9) & "Assets in the month of application: " & intv_app_month_asset & vbCr
objSelection.TypeText chr(9) & "Expenses in the month of application: " & total_expenses & vbCr
objSelection.TypeText chr(9) & chr(9) & "Housing expense in the month of application: " & intv_app_month_housing_expense & vbCr
objSelection.TypeText chr(9) & chr(9) & "Utilities in the month of application: " & utilities_cost & vbCr
If appears_expedited = True Then
	If id_verif_on_file = "No" OR snap_active_in_other_state = "Yes" OR last_snap_was_exp = "Yes" Then
		objSelection.TypeText chr(9) & "Expedited Approval must be delayed:" & vbCr
		objSelection.TypeText chr(9) & chr(9) & "Detail: " & expedited_delay_info & vbCr
		If id_verif_on_file = "No" Then 			objSelection.TypeText chr(9) & chr(9) & "" & vbCr
		If snap_active_in_other_state = "Yes" Then 	objSelection.TypeText chr(9) & chr(9) & "" & vbCr
		If last_snap_was_exp = "Yes" Then 			objSelection.TypeText chr(9) & chr(9) & "" & vbCr
	End If
End If

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
objSelection.TypeParagraph()						'adds a line between the table and the next information

objSelection.Font.Bold = TRUE
objSelection.TypeText "Interview Answers:" & vbCr
objSelection.Font.Bold = FALSE
objSelection.TypeText chr(9) & "Intends to reside in MN? - " & HH_MEMB_ARRAY(intend_to_reside_in_mn, 0) & vbCr
objSelection.TypeText chr(9) & "Has Sponsor? - " & HH_MEMB_ARRAY(clt_has_sponsor, 0) & vbCr
objSelection.TypeText chr(9) & "Immigration Status: " & HH_MEMB_ARRAY(imig_status, 0) & vbCr
objSelection.TypeText chr(9) & "Verification: " & HH_MEMB_ARRAY(client_verification, 0) & vbCr
If HH_MEMB_ARRAY(client_verification_details, 0) <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & HH_MEMB_ARRAY(client_verification_details, 0) & vbCr

'Now we have a dynamic number of tables
'each table has to be defined with its index so we need to have a variable to increment
table_count = 4			'table index variable
If UBound(HH_MEMB_ARRAY, 2) <> 0 Then
	ReDim TABLE_ARRAY(UBound(HH_MEMB_ARRAY, 2)-1)		'defining the table array for as many persons aas are in the household - each person gets their own table
	array_counters = 0		'the incrementer for the table array'

	For each_member = 1 to UBound(HH_MEMB_ARRAY, 2)
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

		TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text = HH_MEMB_ARRAY(last_name_const, each_member)
		TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text = HH_MEMB_ARRAY(first_name_const, each_member)
		TABLE_ARRAY(array_counters).Cell(2, 3).Range.Text = HH_MEMB_ARRAY(mid_initial, each_member)
		TABLE_ARRAY(array_counters).Cell(2, 4).Range.Text = HH_MEMB_ARRAY(other_names, each_member)

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

		TABLE_ARRAY(array_counters).Cell(4, 1).Range.Text = HH_MEMB_ARRAY(ssn, each_member)
		TABLE_ARRAY(array_counters).Cell(4, 2).Range.Text = HH_MEMB_ARRAY(date_of_birth, each_member)
		TABLE_ARRAY(array_counters).Cell(4, 3).Range.Text = HH_MEMB_ARRAY(gender, each_member)
		TABLE_ARRAY(array_counters).Cell(4, 4).Range.Text = HH_MEMB_ARRAY(rel_to_applcnt, each_member)
		TABLE_ARRAY(array_counters).Cell(4, 5).Range.Text = Left(HH_MEMB_ARRAY(marital_status, each_member), 1)

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

		TABLE_ARRAY(array_counters).Cell(6, 1).Range.Text = HH_MEMB_ARRAY(interpreter, each_member)
		TABLE_ARRAY(array_counters).Cell(6, 2).Range.Text = HH_MEMB_ARRAY(spoken_lang, each_member)
		TABLE_ARRAY(array_counters).Cell(6, 3).Range.Text = HH_MEMB_ARRAY(written_lang, each_member)

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

		TABLE_ARRAY(array_counters).Cell(8, 1).Range.Text = HH_MEMB_ARRAY(last_grade_completed, each_member)
		TABLE_ARRAY(array_counters).Cell(8, 2).Range.Text = "Date: " & HH_MEMB_ARRAY(mn_entry_date, each_member) & "   From: " & HH_MEMB_ARRAY(former_state, each_member)
		TABLE_ARRAY(array_counters).Cell(8, 3).Range.Text = HH_MEMB_ARRAY(citizen, each_member)

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
		If HH_MEMB_ARRAY(none_req_checkbox, each_member) = checked then progs_applying_for = "NONE"
		If HH_MEMB_ARRAY(snap_req_checkbox, each_member) = checked then progs_applying_for = progs_applying_for & ", SNAP"
		If HH_MEMB_ARRAY(cash_req_checkbox, each_member) = checked then progs_applying_for = progs_applying_for & ", Cash"
		If HH_MEMB_ARRAY(emer_req_checkbox, each_member) = checked then progs_applying_for = progs_applying_for & ", Emergency Assistance"
		If left(progs_applying_for, 2) = ", " Then progs_applying_for = right(progs_applying_for, len(progs_applying_for) - 2)

		race_to_enter = ""
		If HH_MEMB_ARRAY(race_a_checkbox, each_member) = checked then race_to_enter = race_to_enter & ", Asian"
		If HH_MEMB_ARRAY(race_b_checkbox, each_member) = checked then race_to_enter = race_to_enter & ", Black"
		If HH_MEMB_ARRAY(race_n_checkbox, each_member) = checked then race_to_enter = race_to_enter & ", American Indian or Alaska Native"
		If HH_MEMB_ARRAY(race_p_checkbox, each_member) = checked then race_to_enter = race_to_enter & ", Pacific Islander and Native Hawaiian"
		If HH_MEMB_ARRAY(race_w_checkbox, each_member) = checked then race_to_enter = race_to_enter & ", White"
		If left(race_to_enter, 2) = ", " Then race_to_enter = right(race_to_enter, len(race_to_enter) - 2)

		TABLE_ARRAY(array_counters).Cell(10, 1).Range.Text = progs_applying_for
		TABLE_ARRAY(array_counters).Cell(10, 2).Range.Text = HH_MEMB_ARRAY(ethnicity_yn, each_member)
		TABLE_ARRAY(array_counters).Cell(10, 3).Range.Text = race_to_enter


		objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing

		objSelection.TypeText "INTERVIEW NOTES: " & HH_MEMB_ARRAY(client_notes, each_member) & vbCR
		' objSelection.Font.Bold = TRUE
		' objSelection.TypeText "AGENCY USE:" & vbCr
		' objSelection.Font.Bold = FALSE
		objSelection.TypeText chr(9) & "Intends to reside in MN? - " & HH_MEMB_ARRAY(intend_to_reside_in_mn, each_member) & vbCr
		objSelection.TypeText chr(9) & "Has Sponsor? - " & HH_MEMB_ARRAY(clt_has_sponsor, each_member) & vbCr
		objSelection.TypeText chr(9) & "Immigration Status: " & HH_MEMB_ARRAY(imig_status, each_member) & vbCr
		objSelection.TypeText chr(9) & "Verification: " & HH_MEMB_ARRAY(client_verification, each_member) & vbCr
		If HH_MEMB_ARRAY(client_verification_details, each_member) <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & HH_MEMB_ARRAY(client_verification_details, each_member) & vbCr

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
If question_1_yn <> "" OR trim(question_1_notes) <> "" Then
	objSelection.TypeText chr(9) & "CAF Answer Confirmed during the Interview" & vbCR
	If question_1_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_1_interview_notes & vbCR
End If

objSelection.TypeText "Q 2. Is anyone in the household, who is age 60 or over or disabled, unable to buy or fix food due to a disability?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_2_yn & vbCr
If question_2_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_2_notes & vbCr
If question_2_verif_yn <> "Mot Needed" AND question_2_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_2_verif_yn & vbCr
If question_2_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_2_verif_details & vbCr
If question_2_yn <> "" OR trim(question_2_notes) <> "" Then
	objSelection.TypeText chr(9) & "CAF Answer Confirmed during the Interview" & vbCR
	If question_2_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_2_interview_notes & vbCR
End If

objSelection.TypeText "Q 3. Is anyone in the household attending school?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_3_yn & vbCr
If question_3_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_3_notes & vbCr
If question_3_verif_yn <> "Mot Needed" AND question_3_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_3_verif_yn & vbCr
If question_3_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_3_verif_details & vbCr
If question_3_yn <> "" OR trim(question_3_notes) <> "" Then
	objSelection.TypeText chr(9) & "CAF Answer Confirmed during the Interview" & vbCR
	If question_3_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_3_interview_notes & vbCR
End If

objSelection.TypeText "Q 4. Is anyone in your household temporarily not living in your home? (for example: vacation, foster care, treatment, hospital, job search)" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_4_yn & vbCr
If question_4_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_4_notes & vbCr
If question_4_verif_yn <> "Mot Needed" AND question_4_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_4_verif_yn & vbCr
If question_4_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_4_verif_details & vbCr
If question_4_yn <> "" OR trim(question_4_notes) <> "" Then
	objSelection.TypeText chr(9) & "CAF Answer Confirmed during the Interview" & vbCR
	If question_4_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_4_interview_notes & vbCR
End If

objSelection.TypeText "Q 5. Is anyone blind, or does anyone have a physical or mental health condition that limits the ability to work or perform daily activities?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_5_yn & vbCr
If question_5_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_5_notes & vbCr
If question_5_verif_yn <> "Mot Needed" AND question_5_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_5_verif_yn & vbCr
If question_5_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_5_verif_details & vbCr
If question_5_yn <> "" OR trim(question_5_notes) <> "" Then
	objSelection.TypeText chr(9) & "CAF Answer Confirmed during the Interview" & vbCR
	If question_5_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_5_interview_notes & vbCR
End If

objSelection.TypeText "Q 6. Is anyone unable to work for reasons other than illness or disability?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_6_yn & vbCr
If question_6_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_6_notes & vbCr
If question_6_verif_yn <> "Mot Needed" AND question_6_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_6_verif_yn & vbCr
If question_6_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_6_verif_details & vbCr
If question_6_yn <> "" OR trim(question_6_notes) <> "" Then
	objSelection.TypeText chr(9) & "CAF Answer Confirmed during the Interview" & vbCR
	If question_6_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_6_interview_notes & vbCR
End If

objSelection.TypeText "Q 7. In the last 60 days did anyone in the household: - Stop working or quit a job? - Refuse a job offer? - Ask to work fewer hours? - Go on strike?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_7_yn & vbCr
If question_7_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_7_notes & vbCr
If question_7_verif_yn <> "Mot Needed" AND question_7_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_7_verif_yn & vbCr
If question_7_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_7_verif_details & vbCr
If question_7_yn <> "" OR trim(question_7_notes) <> "" Then
	objSelection.TypeText chr(9) & "CAF Answer Confirmed during the Interview" & vbCR
	If question_7_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_7_interview_notes & vbCR
End If

objSelection.TypeText "Q 8. Has anyone in the household had a job or been self-employed in the past 12 months?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_8_yn & vbCr
objSelection.TypeText "Q 8.a. FOR SNAP ONLY: Has anyone in the household had a job or been self-employed in the past 36 months?" & vbCr
objSelection.TypeText chr(9) & question_8a_yn & vbCr
If question_8_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_8_notes & vbCr
If question_8_verif_yn <> "Mot Needed" AND question_8_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_8_verif_yn & vbCr
If question_8_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_8_verif_details & vbCr
If question_8_yn <> "" OR trim(question_8_notes) <> "" Then
	objSelection.TypeText chr(9) & "CAF Answer Confirmed during the Interview" & vbCR
	If question_8_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_8_interview_notes & vbCR
End If

objSelection.TypeText "Q 9. Does anyone in the household have a job or expect to get income from a job this month or next month?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_9_yn & vbCr
job_added = FALSE
for each_job = 0 to UBOUND(JOBS_ARRAY, 2)
	If JOBS_ARRAY(jobs_employer_name, each_job) <> "" OR JOBS_ARRAY(jobs_employee_name, each_job) <> "" OR JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" OR JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then
		job_added = TRUE

		all_the_tables = UBound(TABLE_ARRAY) + 1
		ReDim Preserve TABLE_ARRAY(all_the_tables)
		Set objRange = objSelection.Range					'range is needed to create tables
		objDoc.Tables.Add objRange, 8, 1					'This sets the rows and columns needed row then column'
		set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
		table_count = table_count + 1

		TABLE_ARRAY(array_counters).AutoFormat(16)							'This adds the borders to the table and formats it
		TABLE_ARRAY(array_counters).Columns(1).Width = 400

		for row = 1 to 7 Step 2
			TABLE_ARRAY(array_counters).Cell(row, 1).SetHeight 10, 2
		Next
		for row = 2 to 8 Step 2
			TABLE_ARRAY(array_counters).Cell(row, 1).SetHeight 15, 2
		Next

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

		TABLE_ARRAY(array_counters).Cell(5, 1).Range.Text = "CAF NOTES"
		TABLE_ARRAY(array_counters).Cell(6, 1).Range.Text = JOBS_ARRAY(jobs_notes, each_job)

		TABLE_ARRAY(array_counters).Cell(7, 1).Range.Text = "INTERVIEW NOTES"
		TABLE_ARRAY(array_counters).Cell(8, 1).Range.Text = JOBS_ARRAY(jobs_intv_notes, each_job)

		objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
		' objSelection.TypeParagraph()						'adds a line between the table and the next information

		array_counters = array_counters + 1

		objSelection.TypeText "Verification: " & JOBS_ARRAY(verif_yn, each_job) & " - " & JOBS_ARRAY(verif_details, each_job) & vbCR
	End If
next

If job_added = FALSE Then objSelection.TypeText chr(9) & "THERE ARE NO JOBS ENTERED." & vbCr

If question_9_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_9_notes & vbCr
' If question_9_verif_yn <> "Mot Needed" AND question_10_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_9_verif_yn & vbCr
' If question_9_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_9_verif_details & vbCr

objSelection.TypeText "Q 10. Is anyone in the household self-employed or does anyone expect to get income from self-employment this month or next month?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_10_yn & vbCr
If question_10_monthly_earnings <> "" Then objSelection.TypeText chr(9) & "Gross Monthly Earnings: " & question_10_monthly_earnings & vbCr
If question_10_monthly_earnings = "" Then objSelection.TypeText chr(9) & "Gross Monthly Earnings: NONE LISTED" & vbCr
If question_10_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_10_notes & vbCr
If question_10_verif_yn <> "Mot Needed" AND question_10_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_10_verif_yn & vbCr
If question_10_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_10_verif_details & vbCr
If question_10_yn <> "" OR trim(question_10_notes) <> "" Then
	objSelection.TypeText chr(9) & "CAF Answer Confirmed during the Interview" & vbCR
	If question_10_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_10_interview_notes & vbCR
End If

objSelection.TypeText "Q 11. Do you expect any changes in income, expenses or work hours?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_11_yn & vbCr
If question_11_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_11_notes & vbCr
If question_11_verif_yn <> "Mot Needed" AND question_11_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_11_verif_yn & vbCr
If question_11_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_11_verif_details & vbCr
If question_11_yn <> "" OR trim(question_11_notes) <> "" Then
	objSelection.TypeText chr(9) & "CAF Answer Confirmed during the Interview" & vbCR
	If question_11_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_11_interview_notes & vbCR
End If

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
q_12_answered = FALSE
If question_12_rsdi_yn <> "" Then q_12_answered = TRUE
If question_12_rsdi_amt <> "" Then q_12_answered = TRUE
If question_12_ssi_yn <> "" Then q_12_answered = TRUE
If question_12_ssi_amt <> "" Then q_12_answered = TRUE
If question_12_va_yn <> "" Then q_12_answered = TRUE
If question_12_va_amt <> "" Then q_12_answered = TRUE
If question_12_ui_yn <> "" Then q_12_answered = TRUE
If question_12_ui_amt <> "" Then q_12_answered = TRUE
If question_12_wc_yn <> "" Then q_12_answered = TRUE
If question_12_wc_amt <> "" Then q_12_answered = TRUE
If question_12_ret_yn <> "" Then q_12_answered = TRUE
If question_12_ret_amt <> "" Then q_12_answered = TRUE
If question_12_trib_yn <> "" Then q_12_answered = TRUE
If question_12_trib_amt <> "" Then q_12_answered = TRUE
If question_12_cs_yn <> "" Then q_12_answered = TRUE
If question_12_cs_amt <> "" Then q_12_answered = TRUE
If question_12_other_yn <> "" Then q_12_answered = TRUE
If question_12_other_amt <> "" Then q_12_answered = TRUE
If question_12_notes <> "" Then q_12_answered = TRUE
If q_12_answered = TRUE  Then
	objSelection.TypeText chr(9) & "CAF Answer Confirmed during the Interview" & vbCR
	If question_12_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_12_interview_notes & vbCR
End If

objSelection.TypeText "Q 13. Does anyone in the household have or expect to get any loans, scholarships or grants for attending school?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_13_yn & vbCr
If question_13_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_13_notes & vbCr
If question_13_verif_yn <> "Mot Needed" AND question_13_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_13_verif_yn & vbCr
If question_13_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_13_verif_details & vbCr
If question_13_yn <> "" OR trim(question_13_notes) <> "" Then
	objSelection.TypeText chr(9) & "CAF Answer Confirmed during the Interview" & vbCR
	If question_13_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_13_interview_notes & vbCR
End If

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
q_14_answered = FALSE
If question_14_rent_yn <> "" Then q_14_answered = TRUE
If question_14_subsidy_yn <> "" Then q_14_answered = TRUE
If question_14_mortgage_yn <> "" Then q_14_answered = TRUE
If question_14_association_yn <> "" Then q_14_answered = TRUE
If question_14_insurance_yn <> "" Then q_14_answered = TRUE
If question_14_room_yn <> "" Then q_14_answered = TRUE
If question_14_taxes_yn <> "" Then q_14_answered = TRUE
If q_14_answered = TRUE  Then
	objSelection.TypeText chr(9) & "CAF Answer Confirmed during the Interview" & vbCR
	If question_14_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_14_interview_notes & vbCR
End If

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
q_15_answered = FALSE
If question_15_heat_ac_yn <> "" Then q_15_answered = TRUE
If question_15_electricity_yn <> "" Then q_15_answered = TRUE
If question_15_cooking_fuel_yn <> "" Then q_15_answered = TRUE
If question_15_water_and_sewer_yn <> "" Then q_15_answered = TRUE
If question_15_garbage_yn <> "" Then q_15_answered = TRUE
If question_15_phone_yn <> "" Then q_15_answered = TRUE
If question_15_liheap_yn <> "" Then q_15_answered = TRUE
If q_15_answered = TRUE  Then
	objSelection.TypeText chr(9) & "CAF Answer Confirmed during the Interview" & vbCR
	If question_15_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_15_interview_notes & vbCR
End If

objSelection.TypeText "Q 16. Do you or anyone living with you have costs for care of a child(ren) because you or they are working, looking for work or going to school?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_16_yn & vbCr
If question_16_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_16_notes & vbCr
If question_16_verif_yn <> "Mot Needed" AND question_16_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_16_verif_yn & vbCr
If question_16_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_16_verif_details & vbCr
If question_16_yn <> "" OR trim(question_16_notes) <> "" Then
	objSelection.TypeText chr(9) & "CAF Answer Confirmed during the Interview" & vbCR
	If question_16_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_16_interview_notes & vbCR
End If

objSelection.TypeText "Q 17. Do you or anyone living with you have costs for care of an ill or disabled adult because you or they are working, looking for work or going to school?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_17_yn & vbCr
If question_17_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_17_notes & vbCr
If question_17_verif_yn <> "Mot Needed" AND question_17_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_17_verif_yn & vbCr
If question_17_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_17_verif_details & vbCr
If question_17_yn <> "" OR trim(question_17_notes) <> "" Then
	objSelection.TypeText chr(9) & "CAF Answer Confirmed during the Interview" & vbCR
	If question_17_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_17_interview_notes & vbCR
End If

objSelection.TypeText "Q 18. Does anyone in the household pay court-ordered child support, spousal support, child care support, medical support or contribute to a tax dependent who does not live in your home?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_18_yn & vbCr
If question_18_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_18_notes & vbCr
If question_18_verif_yn <> "Mot Needed" AND question_18_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_18_verif_yn & vbCr
If question_18_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_18_verif_details & vbCr
If question_18_yn <> "" OR trim(question_18_notes) <> "" Then
	objSelection.TypeText chr(9) & "CAF Answer Confirmed during the Interview" & vbCR
	If question_18_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_18_interview_notes & vbCR
End If

objSelection.TypeText "Q 19. For SNAP only: Does anyone in the household have medical expenses? " & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_19_yn & vbCr
If question_19_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_19_notes & vbCr
If question_19_verif_yn <> "Mot Needed" AND question_19_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_19_verif_yn & vbCr
If question_19_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_19_verif_details & vbCr
If question_19_yn <> "" OR trim(question_19_notes) <> "" Then
	objSelection.TypeText chr(9) & "CAF Answer Confirmed during the Interview" & vbCR
	If question_19_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_19_interview_notes & vbCR
End If

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
q_20_answered = FALSE
If question_20_cash_yn <> "" Then q_20_answered = TRUE
If question_20_acct_yn <> "" Then q_20_answered = TRUE
If question_20_secu_yn <> "" Then q_20_answered = TRUE
If question_20_cars_yn <> "" Then q_20_answered = TRUE
If q_20_answered = TRUE  Then
	objSelection.TypeText chr(9) & "CAF Answer Confirmed during the Interview" & vbCR
	If question_20_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_20_interview_notes & vbCR
End If

objSelection.TypeText "Q 21. For Cash programs only: Has anyone in the household given away, sold or traded anything of value in the past 12 months? (For example: Cash, Bank accounts, Stocks, Bonds, Vehicles)" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_21_yn & vbCr
If question_21_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_21_notes & vbCr
If question_21_verif_yn <> "Mot Needed" AND question_21_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_21_verif_yn & vbCr
If question_21_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_21_verif_details & vbCr
If question_21_yn <> "" OR trim(question_21_notes) <> "" Then
	objSelection.TypeText chr(9) & "CAF Answer Confirmed during the Interview" & vbCR
	If question_21_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_21_interview_notes & vbCR
End If

objSelection.TypeText "Q 22. For recertifications only: Did anyone move in or out of your home in the past 12 months?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_22_yn & vbCr
If question_22_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_22_notes & vbCr
If question_22_verif_yn <> "Mot Needed" AND question_22_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_22_verif_yn & vbCr
If question_22_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_22_verif_details & vbCr
If question_22_yn <> "" OR trim(question_22_notes) <> "" Then
	objSelection.TypeText chr(9) & "CAF Answer Confirmed during the Interview" & vbCR
	If question_22_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_22_interview_notes & vbCR
End If

objSelection.TypeText "Q 23. For children under the age of 19, are both parents living in the home?" & vbCr
objSelection.TypeText chr(9) & "CAF Answer: " & question_23_yn & vbCr
If question_23_notes <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & question_23_notes & vbCr
If question_23_verif_yn <> "Mot Needed" AND question_23_verif_yn <> "" Then objSelection.TypeText chr(9) & "Verification: " & question_23_verif_yn & vbCr
If question_23_verif_details <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & question_23_verif_details & vbCr
If question_23_yn <> "" OR trim(question_23_notes) <> "" Then
	objSelection.TypeText chr(9) & "CAF Answer Confirmed during the Interview" & vbCR
	If question_23_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_23_interview_notes & vbCR
End If

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
q_24_answered = FALSE
If question_24_rep_payee_yn <> "" Then q_24_answered = TRUE
If question_24_guardian_fees_yn <> "" Then q_24_answered = TRUE
If question_24_special_diet_yn <> "" Then q_24_answered = TRUE
If question_24_high_housing_yn <> "" Then q_24_answered = TRUE
If q_24_answered = TRUE  Then
	objSelection.TypeText chr(9) & "CAF Answer Confirmed during the Interview" & vbCR
	If question_24_interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & question_24_interview_notes & vbCR
End If

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

	' 'Needs to determine MyDocs directory before proceeding.
	' Set wshshell = CreateObject("WScript.Shell")
	' user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"
	' 'this is the file for the 'save your work' functionality.
	' If MAXIS_case_number <> "" Then
	' 	local_changelog_path = user_myDocs_folder & "caf-answers-" & MAXIS_case_number & "-info.txt"
	' Else
	' 	local_changelog_path = user_myDocs_folder & "caf-answers-new-case-info.txt"
	' End If
	'
	' 'we are checking the save your work text file. If it exists we need to delete it because we don't want to save that information locally.
	' If objFSO.FileExists(local_changelog_path) = True then
	' 	objFSO.DeleteFile(local_changelog_path)			'DELETE
	' End If

	' 'Now we case note!
	Call start_a_blank_case_note
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
	end_msg = "Success! The information you have provided about the interview and all of the notes have been saved in a PDF. This PDF will be uploaded to ECF by SSR staff for Case # " & MAXIS_case_number & " and will remain in the CASE RECORD. CASE:NOTES have also been entered with the full interview detail."

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



'POLICY NOTES
'
' Here is what Ann from Internal Services said about additional training:
'
' There is a training in IPAM that covers how to interview and covers annotating.
'
' Per CM
' WHAT IS A COMPLETE APPLICATION (state.mn.us)
' https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=cm_00051203
' obtain the answers from the client at the time of the interview and clearly document the information provided.
'
' APPLICATION INTERVIEWS (state.mn.us)
' https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=cm_00051212
' Nothing mentioned in this section either
'
' IPAM
' An Eligibility Workers Guide to the Combined Application Form (With Answers).pdf (state.mn.us)
' https://www.dhssir.cty.dhs.state.mn.us/MAXIS/trntl/_layouts/15/WopiFrame.aspx?sourcedoc=%7B3230AF4F-4FA7-448C-BAA7-506671E03A49%7D&file=An%20Eligibility%20Workers%20Guide%20to%20the%20Combined%20Application%20Form%20(With%20Answers).pdf&action=default&IsList=1&ListId=%7B032C9304-E9F4-4ED6-90A0-92F9CC18CD31%7D&ListItemId=2
' Answer section page 64
' 1) On what form do you record information from the interview?
' Information from the interview must be recorded on the CAF and in MAXIS CASE/NOTES, in sufficient detail for other workers and supervisors to follow the adequacy of the certification process and the accuracy of your decisions.
