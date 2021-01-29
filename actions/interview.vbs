'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - INTERVIEW.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 0                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
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
call changelog_update("07/13/2020", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

BeginDialog Dialog1, 0, 0, 550, 385, "Dialog"
  GroupBox 180, 5, 300, 260, "Client Conversation"
  GroupBox 10, 5, 170, 260, "Address Information listed in MAXIS"
  ButtonGroup ButtonPressed
    PushButton 50, 245, 125, 15, "CAF has Different Information", caf_info_different_btn
    PushButton 485, 10, 60, 15, "CAF Page 1", caf_page_one_btn
    PushButton 415, 365, 50, 15, "NEXT", next_btn
    PushButton 465, 365, 80, 15, "Complete Interview", finish_interview_btn
EndDialog

function read_and_format_from_MAXIS()
end function

class mx_hh_member

	public access_denied
	public selected
	'stuff about the members
	public first_name
	public last_name
	public mid_initial
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
	public snap_minor
	public cash_minor
	public written_lang
	public spoken_lang

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

	public spon_exists
	public clt_has_sponsor
	public ask_about_spon
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
	public spon_possible_indigent_exemption
	public spon_gross_income
	public spon_spouse_income


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
			End If

			Call navigate_to_MAXIS_screen("STAT", "SPON")		'===============================================================================================
			EMWriteScreen ref_number, 20, 76
			transmit

			EMreadScreen spon_version, 1, 2, 73
			If spon_version = "0" Then spon_exists = FALSE
			If spon_version = "1" Then spon_exists = TRUE

			If spon_exists = TRUE Then


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

	public sub assess_imig_questions()
		imig_q_2_required = FALSE
		If left(imig_status, 2) = "24" Then imig_q_2_required = TRUE

		ask_about_spon = TRUE
		If left(imig_status, 2) = "21" Then ask_about_spon = FALSE
		If left(imig_status, 2) = "22" Then ask_about_spon = FALSE
		If left(imig_status, 2) = "28" Then ask_about_spon = FALSE
		If left(imig_prev_status, 2) = "21" Then ask_about_spon = FALSE
		If left(imig_prev_status, 2) = "22" Then ask_about_spon = FALSE

		imig_q_4_required = FALSE
		imig_q_5_required = FALSE
		If clt_has_sponsor = "Yes" Then
			imig_q_4_required = TRUE
			imig_q_5_required = TRUE
		End If

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


Dim HH_MEMB_ARRAY()
ReDim HH_MEMB_ARRAY(0)

Dim INCOME_ARRAY()
ReDim INCOME_ARRAY(0)

const rela_clt_one_ref		= 0
const rela_clt_two_ref 		= 1
const rela_clt_one_name 	= 2
const rela_clt_two_name 	= 3
const rela_type 			= 4
const verif_req_checkbox 	= 5
const rela_verif			= 6
const rela_pers_one			= 7
const rela_pers_two			= 8
const rela_notes 			= 9

Dim ALL_HH_RELATIONSHIPS_ARRAY()
ReDim ALL_HH_RELATIONSHIPS_ARRAY(rela_notes, 0)

const temp_abs_person		= 0
const temp_abs_ref			= 1
const temp_abs_name			= 2
const temp_abs_where		= 3
const temp_abs_left_date	= 4
const temp_abs_ret_date		= 5
const temp_abs_notes		= 6

Dim ALL_TEMP_ABSENCE()
ReDim ALL_TEMP_ABSENCE(temp_abs_notes, 0)

const pers_unable_to_work 		= 0
const ref_nbr_unable_to_work 	= 1
const name_unable_to_work 		= 2
const unable_to_work_start_date	= 3
const unable_to_work_verif 		= 4
const unable_to_work_reason 	= 5
const unable_to_work_abawd_yn	= 6
const unable_to_work_abawd_type	= 7
const unable_to_work_mfip_yn	= 8
const unable_to_work_mfip_type	= 9
const unable_to_work_notes 		= 10

Dim NON_DISA_UNABLE_TO_WORK()
ReDim NON_DISA_UNABLE_TO_WORK(unable_to_work_notes, 0)

const quest_numb = 0
const memb_numb = 1

Dim QUAL_QUESTIONS_ARRAY()
ReDim QUAL_QUESTIONS_ARRAY(memb_numb, 0)

rela_type_dropdown = "Select One..."+chr(9)+"Parent"+chr(9)+"Child"+chr(9)+"Sibling"+chr(9)+"Spouse"+chr(9)+"Grandparent"+chr(9)+"Neice"+chr(9)+"Nephew"+chr(9)+"Aunt"+chr(9)+"Uncle"+chr(9)+"Grandchild"+chr(9)+"Step Parent"+chr(9)+"Step Child"+chr(9)+"Relative Caregiver"+chr(9)+"Foster Child"+chr(9)+"Foster Parent"+chr(9)+"Not Related"+chr(9)+"Legal Guardian"+chr(9)+"Other Relative"+chr(9)+"Cousin"+chr(9)+"Live-in Attendant"+chr(9)+"Unknown"
rela_verif_dropdown = "Type or Select"+chr(9)+"BC - Birth Certificate"+chr(9)+"AR - Adoption Records"+chr(9)+"LG = Legal Guardian"+chr(9)+"RE - Religious Records"+chr(9)+"HR - Hospital Records"+chr(9)+"RP - Recognition of Parentage"+chr(9)+"OT - Other Verifciation"+chr(9)+"NO - No Verif Provided"+chr(9)
grade_droplist = "Select One..."+chr(9)+"Kindergarten"+chr(9)+"1st Grade"+chr(9)+"2nd Grade"+chr(9)+"3rd Grade"+chr(9)+"4th Grade"+chr(9)+"5th Grade"+chr(9)+"6th Grade"+chr(9)+"7th Grade"+chr(9)+"8th Grade"+chr(9)+"9th Grade"+chr(9)+"10th Grade"+chr(9)+"11th Grade"+chr(9)+"12th Grade"
schl_ver_droplist = "Type or Select"+chr(9)+"Not Needed"+chr(9)+"Requested"+chr(9)+"Received - "+chr(9)+"On File - "+chr(9)+"SC - School Statement"+chr(9)+"OT - Other Document"+chr(9)+"No - No Verif Provided"
schl_status_droplist = "Select One..."+chr(9)+"Fulltime"+chr(9)+"Halftime"+chr(9)+"Less than Half"+chr(9)+"Not Attending"
caf_answer_droplist = " "+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Blank"
unea_verif_droplist = "Select One..."+chr(9)+"1 - Copy of Checks"+chr(9)+"2 - Award Letters"+chr(9)+"3 - System Initiated"+chr(9)+"4 - Colateral Statement"+chr(9)+"5 - Pend Out State Verif"+chr(9)+"6 - Other Document"+chr(9)+"7 - Worker Initiated"+chr(9)+"8 - RI Stubs"+chr(9)+"N - No Verif Provided"+chr(9)
days_of_the_week_droplist = "Select One..."+chr(9)+"Monday"+chr(9)+"Tuesday"+chr(9)+"Wednesday"+chr(9)+"Thursday"+chr(9)+"Friday"+chr(9)+"Saturday"+chr(9)+"Sunday"
imig_status_dropdown = "21 - Refugee"
imig_status_dropdown = imig_status_dropdown+chr(9)+"22 - Asylee"
imig_status_dropdown = imig_status_dropdown+chr(9)+"23 - Deportation/Removal Withheld"
imig_status_dropdown = imig_status_dropdown+chr(9)+"24 - LPR"
imig_status_dropdown = imig_status_dropdown+chr(9)+"25 - Paroled for 1 Year or More"
imig_status_dropdown = imig_status_dropdown+chr(9)+"26 - Conditional Entry < 4/80"
imig_status_dropdown = imig_status_dropdown+chr(9)+"27 - Non-Immmigraant"
imig_status_dropdown = imig_status_dropdown+chr(9)+"28 - Unndocumented"
imig_status_dropdown = imig_status_dropdown+chr(9)+"50 - Other Lawfully Residing"
imig_nationality_droplist = "AA - Amerasian"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"EH - Ethnic Chinese"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"EL - Ethnic Lao"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"HG - Hmong"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"KD - Kurd"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"SJ - Soviet Jew"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"TT - Tinh"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"AF - Afghanistan"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"BK - Bosnia"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"CB - Cambodia"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"CH - China, Mainland"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"CU - Cuba"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"ES - El Salvador"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"ER - Eritrea"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"ET - Ethiopia"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"GT - Guatemala"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"HA - Haiti"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"HO - Honduras"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"IR - Iran"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"IZ - Iraq"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"LI - Liberia"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"MC - Micronesia"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"MI - Marshall Islands"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"MX - Mexico"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"WA - Namibia"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"PK - Pakistan"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"RP - Philippines"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"PL - Poland"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"RO - Romania"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"RS - Russia"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"SO - Somalia"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"SF - South Africa"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"TH - Thailand"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"VM - Vietnam"
imig_nationality_droplist = imig_nationality_droplist+chr(9)+"OT - All Others"
memb_droplist = ""

the_pwe_for_this_case = ""
child_on_case = FALSE

rsdi_count = 0
ssi_count = 0
va_count = 0
ui_count = 0
wc_count = 0
retirement_count = 0
tribal_count = 0
cs_count = 0
ss_count = 0
other_UNEA_count = 0
cash_count = 0
acct_count = 0
secu_count = 0
cars_count = 0
rest_count = 0
memb_to_match = ""
inst_to_match = ""
unea_found = FALSE
jobs_found = FALSE
busi_found = FALSE
cash_found = FALSE
acct_found = FALSE
secu_found = FALSE
cars_found = FALSE
rest_found = FALSE
imig_questions_step = 1

'Button Definitions
caf_page_one_btn	= 1000
caf_membs_btn		= 1001
caf_q_1_2_btn		= 1002
caf_q_3_btn			= 1003
caf_q_4_btn 		= 1004
caf_q_5_btn			= 1005
caf_q_6_btn			= 1006
caf_q_7_btn			= 1007
caf_q_8_btn			= 1008
caf_q_9_btn			= 1009
caf_q_10_btn		= 1010
caf_q_11_btn		= 1011
caf_q_12_btn		= 1012
caf_q_13_btn		= 1013
caf_q_14_15_btn		= 1014
caf_q_16_17_18_btn	= 1015
caf_q_19_btn		= 1016
caf_q_20_21_btn		= 1017
caf_q_22_btn		= 1018
caf_q_23_btn		= 1019
caf_q_24_btn		= 1020
caf_qual_q_btn		= 1021
caf_last_page_btn	= 1022
next_btn			= 1023
finish_interview_btn= 1024

client_verbal_changes_resi_address_btn	= 2000
client_verbal_changes_mail_address_btn	= 2001
caf_info_different_btn					= 2002
client_verbal_changes_phone_numbers_btn	= 2003
add_memb_to_list_btn					= 2004
HH_memb_detail_review					= 2005
add_relationship_btn					= 2006
memb_info_change						= 2007
next_memb_btn							= 2008
hh_list_btn								= 2009
hh_imig_btn								= 2010
update_imig_info_btn					= 2011
update_groups_btn						= 2012
search_district_btn						= 2013
add_higher_ed_studen					= 2014
add_ged_ell_student						= 2015
add_another_absent_pers_btn				= 2016
new_disa_btn							= 2017
add_unable_tp_work_memb_btn				= 2018
next_stwk_btn							= 2019
prev_imig_btn							= 2020
next_imig_btn							= 2021

add_qual_quest_yes						= 2030

rsdi_btn 	= 3000
ssi_btn		= 3001
va_btn		= 3002
ui_btn		= 3003
wc_btn		= 3004
ret_btn		= 3005
tribal_btn	= 3006
cs_btn		= 3007
ss_btn		= 3008
other_btn	= 3009
main_btn	= 3010
cash_btn	= 3011
acct_btn	= 3012
secu_btn	= 3013
cars_btn	= 3014
rest_btn	= 3015
main_asset_btn	= 3016

'PRESETS FOR QUESTIONS COMPLETIONS
done_pg_one 	= FALSE
done_pg_memb 	= FALSE
done_q_1_2		= FALSE
done_q_3		= FALSE
done_q_4		= FALSE
done_q_5		= FALSE
done_q_6		= FALSE
done_q_7		= FALSE
done_q_8		= FALSE
done_q_9		= FALSE
done_q_10		= FALSE
done_q_11		= FALSE
done_q_12		= FALSE
done_q_13		= FALSE
done_q_14_15	= FALSE
done_q_16_18	= FALSE
done_q_19		= FALSE
done_q_20_21 	= FALSE
done_q_22		= FALSE
done_q_23		= FALSE
done_q_24		= FALSE
done_qual		= FALSE
done_pg_last	= FALSE

'SETTINGS FOR PAGE IDETIFIERS
show_pg_one 		= 1
show_pg_memb_list 	= 2
show_pg_memb_info 	= 3
show_pg_imig		= 4
show_q_1_2			= 5
show_q_3			= 6
show_q_4			= 7
show_q_5			= 8
show_q_6			= 9
show_q_7			= 10
show_q_8			= 11
show_q_9			= 12
show_q_10			= 13
show_q_11			= 14
show_q_12			= 15
show_q_13			= 16
show_q_14_15		= 17
show_q_16_18		= 18
show_q_19			= 19
show_q_20_21 		= 20
show_q_22			= 21
show_q_23			= 22
show_q_24			= 23
show_qual			= 24
show_pg_last		= 25

rsdi_unea	= 1
ssi_unea	= 2
va_unea		= 3
ui_unea		= 4
wc_unea		= 5
ret_unea	= 6
tribal_unea	= 7
cs_unea		= 8
ss_unea		= 9
other_unea	= 10
main_unea	= 11
cash_asset	= 12
acct_asset	= 13
secu_asset	= 14
cars_asset	= 15
main_asset	= 16

' page__step__btn 	= 10,000
page_1_step_1_btn 	= 10011
page_1_step_2_btn 	= 10012
page_1_step_3_btn 	= 10013
page_1_step_4_btn 	= 10014
page_1_step_5_btn 	= 10015
page_1_step_6_btn 	= 10016
page_1_step_7_btn 	= 10017
page_1_step_8_btn 	= 10018

Dim case_has_imig

function dialog_movement()
	case_has_imig = FALSE
	For i = 0 to Ubound(HH_MEMB_ARRAY, 1)
		If HH_MEMB_ARRAY(i).imig_exists = TRUE Then case_has_imig = TRUE
		' MsgBox HH_MEMB_ARRAY(i).button_one
		If ButtonPressed = HH_MEMB_ARRAY(i).button_one Then
			If page_display = show_pg_memb_info Then memb_selected = i
			If page_display = show_pg_imig Then memb_selected = i
			If page_display = show_q_12 Then memb_to_match = HH_MEMB_ARRAY(i).ref_number
			If page_display = show_q_7 Then stwk_selected = i
			If page_display = show_q_19 Then fmed_selected = i
			' If second_page_display = ssi_unea Then memb_to_match = HH_MEMB_ARRAY(i).ref_number
		End If
	Next
	For i = 0 to Ubound(ASSET_ARRAY, 1)
		' MsgBox HH_MEMB_ARRAY(i).button_one
		If ButtonPressed = ASSET_ARRAY(i).asset_btn_one Then
			memb_to_match = ASSET_ARRAY(i).member_ref
			inst_to_match = ASSET_ARRAY(i).panel_instance
		End If
	Next
	' MsgBox ButtonPressed
	If page_display = show_pg_memb_info AND ButtonPressed = -1 Then ButtonPressed = next_memb_btn
	If page_display = show_pg_imig AND ButtonPressed = -1 Then ButtonPressed = next_imig_btn

	If ButtonPressed = next_imig_btn Then
		imig_questions_step = imig_questions_step + 1
		HH_MEMB_ARRAY(memb_selected).assess_imig_questions
		If imig_questions_step = 2 AND HH_MEMB_ARRAY(memb_selected).imig_q_2_required = FALSE Then imig_questions_step = 3
		If imig_questions_step = 4 AND HH_MEMB_ARRAY(memb_selected).imig_q_4_required = FALSE Then imig_questions_step = 5
		If imig_questions_step = 5 AND HH_MEMB_ARRAY(memb_selected).imig_q_5_required = FALSE Then imig_questions_step = 6
		If imig_questions_step > 5 Then ButtonPressed = next_memb_btn
	End If
	If ButtonPressed = next_memb_btn Then
		memb_selected = memb_selected + 1
		If memb_selected > UBound(HH_MEMB_ARRAY, 1) Then ButtonPressed = next_btn
	End If
	If ButtonPressed = next_stwk_btn Then
		stwk_selected = stwk_selected + 1
		Do
			If HH_MEMB_ARRAY(stwk_selected).stwk_exists = FALSE Then stwk_selected = stwk_selected + 1
			If stwk_selected > UBound(HH_MEMB_ARRAY, 1) Then Exit Do
		Loop Until  HH_MEMB_ARRAY(stwk_selected).stwk_exists = TRUE
		If stwk_selected > UBound(HH_MEMB_ARRAY, 1) Then ButtonPressed = next_btn
	End If
	If ButtonPressed = next_fmed_btn Then
		fmed_selected = fmed_selected + 1
		Do
			If HH_MEMB_ARRAY(fmed_selected).fmed_exists = FALSE Then fmed_selected = fmed_selected + 1
			If fmed_selected > UBound(HH_MEMB_ARRAY, 1) Then Exit Do
		Loop Until  HH_MEMB_ARRAY(fmed_selected).fmed_exists = TRUE
		If fmed_selected > UBound(HH_MEMB_ARRAY, 1) Then ButtonPressed = next_btn
	End If
	If ButtonPressed = -1 Then ButtonPressed = next_btn
	If ButtonPressed = next_btn Then
		If page_display = show_pg_one Then ButtonPressed = caf_membs_btn
		If page_display = show_pg_memb_list Then ButtonPressed = HH_memb_detail_review
		If page_display = show_pg_memb_info AND case_has_imig = FALSE Then ButtonPressed = show_pg_imig
		If page_display = show_pg_memb_info AND case_has_imig = TRUE Then ButtonPressed = caf_q_1_2_btn
		If ButtonPressed = show_pg_imig Then ButtonPressed = caf_q_1_2_btn
		If page_display = show_q_1_2 Then ButtonPressed = caf_q_3_btn
		If page_display = show_q_3 Then ButtonPressed = caf_q_4_btn
		If page_display = show_q_4 Then ButtonPressed = caf_q_5_btn
		If page_display = show_q_5 Then ButtonPressed = caf_q_6_btn
		If page_display = show_q_6 Then ButtonPressed = caf_q_7_btn
		If page_display = show_q_7 Then ButtonPressed = caf_q_8_btn
		If page_display = show_q_8 Then ButtonPressed = caf_q_9_btn
		If page_display = show_q_9 Then ButtonPressed = caf_q_10_btn
		If page_display = show_q_10 Then ButtonPressed = caf_q_11_btn
		If page_display = show_q_11 Then ButtonPressed = caf_q_12_btn
		If page_display = show_q_12 Then
			If second_page_display = main_unea Then ButtonPressed = rsdi_btn
			If second_page_display = rsdi_unea Then ButtonPressed = ssi_btn
			If second_page_display = ssi_unea Then ButtonPressed = va_btn
			If second_page_display = va_unea Then ButtonPressed = ui_btn
			If second_page_display = ui_unea Then ButtonPressed = wc_btn
			If second_page_display = wc_unea Then ButtonPressed = ret_btn
			If second_page_display = ret_unea Then ButtonPressed = tribal_btn
			If second_page_display = tribal_unea Then ButtonPressed = cs_btn
			If second_page_display = cs_unea Then ButtonPressed = ss_btn
			If second_page_display = ss_unea Then ButtonPressed = other_btn
			If second_page_display = other_unea Then ButtonPressed = caf_q_13_btn
		End If
		If page_display = show_q_13 Then ButtonPressed = caf_q_14_15_btn
		If page_display = show_q_14_15 Then ButtonPressed = caf_q_16_17_18_btn
		If page_display = show_q_16_18 Then ButtonPressed = caf_q_19_btn
		If page_display = show_q_19 Then ButtonPressed = caf_q_20_21_btn
		If page_display = show_q_20_21 Then
			If second_page_display = main_asset Then ButtonPressed = cash_btn
			If second_page_display = cash_asset Then ButtonPressed = acct_btn
			If second_page_display = acct_asset Then ButtonPressed = secu_btn
			If second_page_display = secu_asset Then ButtonPressed = cars_btn
			If second_page_display = cars_asset Then ButtonPressed = rest_btn
			If second_page_display = rest_asset Then ButtonPressed = caf_q_22_btn
		End If
		If page_display = show_q_22 Then ButtonPressed = caf_q_23_btn
		If page_display = show_q_23 Then ButtonPressed = caf_q_24_btn
		If page_display = show_q_24 Then ButtonPressed = caf_qual_q_btn
		If page_display = show_qual Then ButtonPressed = caf_last_page_btn
		' If page_display = show_pg_last Then ButtonPressed =
	End If

	If ButtonPressed = caf_page_one_btn Then
		page_display = show_pg_one
	End If
	If ButtonPressed = caf_membs_btn Then
		page_display = show_pg_memb_list
	End If
	If ButtonPressed = hh_list_btn Then
		page_display = show_pg_memb_list
	End If
	If ButtonPressed = HH_memb_detail_review Then
		page_display = show_pg_memb_info
	End If
	If ButtonPressed = hh_imig_btn Then
		page_display = show_pg_imig
	End If
	If ButtonPressed = caf_q_1_2_btn Then
		page_display = show_q_1_2
	End If
	If ButtonPressed = caf_q_3_btn Then
		page_display = show_q_3
	End If
	If ButtonPressed = caf_q_4_btn Then
		page_display = show_q_4
	End If
	If ButtonPressed = caf_q_5_btn Then
		page_display = show_q_5
	End If
	If ButtonPressed = caf_q_6_btn Then
		page_display = show_q_6
	End If
	If ButtonPressed = caf_q_7_btn Then
		page_display = show_q_7
	End If
	If ButtonPressed = caf_q_8_btn Then
		page_display = show_q_8
	End If
	If ButtonPressed = caf_q_9_btn Then
		page_display = show_q_9
	End If
	If ButtonPressed = caf_q_10_btn Then
		page_display = show_q_10
	End If
	If ButtonPressed = caf_q_11_btn Then
		page_display = show_q_11
	End If
	If ButtonPressed = caf_q_12_btn Then
		page_display = show_q_12
		second_page_display = main_unea
	End If
	If ButtonPressed = rsdi_btn 	Then second_page_display = rsdi_unea
	If ButtonPressed = ssi_btn		Then second_page_display = ssi_unea
	If ButtonPressed = va_btn		Then second_page_display = va_unea
	If ButtonPressed = ui_btn		Then second_page_display = ui_unea
	If ButtonPressed = wc_btn		Then second_page_display = wc_unea
	If ButtonPressed = ret_btn		Then second_page_display = ret_unea
	If ButtonPressed = tribal_btn	Then second_page_display = tribal_unea
	If ButtonPressed = cs_btn		Then second_page_display = cs_unea
	If ButtonPressed = ss_btn		Then second_page_display = ss_unea
	If ButtonPressed = other_btn	Then second_page_display = other_unea
	If ButtonPressed = main_btn		Then second_page_display = main_unea

	If ButtonPressed = caf_q_13_btn Then
		page_display = show_q_13
	End If
	If ButtonPressed = caf_q_14_15_btn Then
		page_display = show_q_14_15
	End If
	If ButtonPressed = caf_q_16_17_18_btn Then
		page_display = show_q_16_18
	End If
	If ButtonPressed = caf_q_19_btn Then
		page_display = show_q_19
	End If
	If ButtonPressed = caf_q_20_21_btn Then
		page_display = show_q_20_21
		second_page_display = main_asset
	End If
	If ButtonPressed = cash_btn		Then second_page_display = cash_asset
	If ButtonPressed = acct_btn		Then second_page_display = acct_asset
	If ButtonPressed = secu_btn		Then second_page_display = secu_asset
	If ButtonPressed = cars_btn		Then second_page_display = cars_asset
	If ButtonPressed = rest_btn		Then second_page_display = rest_asset
	If ButtonPressed = main_asset_btn		Then second_page_display = main_asset
	If ButtonPressed = caf_q_22_btn Then
		page_display = show_q_22
	End If
	If ButtonPressed = caf_q_23_btn Then
		page_display = show_q_23
	End If
	If ButtonPressed = caf_q_24_btn Then
		page_display = show_q_24
	End If
	If ButtonPressed = caf_qual_q_btn Then
		page_display = show_qual
	End If
	If ButtonPressed = caf_last_page_btn Then
		page_display = show_pg_last
	End If
	If ButtonPressed = finish_interview_btn then leave_loop = TRUE

	If page_display <> show_pg_memb_info AND page_display <> show_pg_imig Then memb_selected = ""
	If page_display <> show_q_7 Then stwk_selected = ""
	If page_display <> show_q_12 AND page_display <> show_q_20_21 Then memb_to_match = ""
	If page_display <> show_q_19 Then fmed_selected = ""

	If page_display <> show_q_20_21 Then inst_to_match = ""

	If ButtonPressed > 10000 Then
		save_button = ButtonPressed
		If ButtonPressed = page_1_step_1_btn Then call explain_dialog_actions("PAGE 1", "STEP 1")
		If ButtonPressed = page_1_step_2_btn Then call explain_dialog_actions("PAGE 1", "STEP 2")

		ButtonPressed = save_button
	End If

end function

function explain_dialog_actions(dlg_page, dlg_step)
	BeginDialog Dialog1, 0, 0, 550, 385, "Full Interview Questions"
		ButtonGroup ButtonPressed
		If dlg_page = "PAGE 1" Then
			Text 495, 12, 60, 13, "CAF Page 1"

			If dlg_step = "STEP 1" Then
				GroupBox 10, 5, 170, 260, "Address Information listed in MAXIS"
				Text 15, 20, 160, 25, "^^1 - Compare this address to the one entered on the CAF. If there is a difference, press the button below to update."
				Text 20, 60, 70, 10, "Residence Address"
				Text 25, 75, 150, 10, resi_line_one
				If resi_line_two = "" Then
					Text 25, 85, 150, 10, resi_city & ", " & resi_state & " " & resi_zip
					Text 25, 95, 150, 10, "County: " & resi_county

				Else
					Text 25, 85, 150, 10, resi_line_two
					Text 25, 95, 150, 10, resi_city & ", " & resi_state & " " & resi_zip
					Text 25, 105, 150, 10, "County: " & resi_county
				End If
				Text 25, 120, 65, 10, "Homeless: " & homeless
				Text 25, 130, 150, 20, "Living Situation: " & living_sit
				Text 25, 155, 60, 10, "IND RES - " & ind_reservation
				Text 25, 165, 150, 10, "RES NAME - " & res_name
				Text 25, 180, 150, 10, "Verification: " & verif
				If mail_line_one = "" Then
					Text 20, 195, 150, 10, "No MAILING ADDRESS Listed"
				ElseIf mail_line_two = "" Then
					Text 20, 195, 70, 10, "Mailing Address"
					Text 25, 210, 150, 10, mail_line_one
					Text 25, 220, 150, 10, mail_city & ", " & mail_state & " " & mail_zip
				Else
					Text 20, 195, 70, 10, "Mailing Address"
					Text 25, 210, 150, 10, mail_line_one
					Text 25, 220, 150, 10, mail_line_two
					Text 25, 230, 150, 10, mail_city & ", " & mail_state & " " & mail_zip
				End If
				' PushButton 50, 245, 125, 15, "CAF has Different Information", caf_info_different_btn
				GroupBox 50, 242, 125, 20, ""
				text 67, 250, 100, 10, "CAF has Different Information"

				Text 185, 20, 20, 10, "<<---"
				Text 200, 20, 200, 25, "This is asking you to check the CAF form received on this case and see if there has been any change reported in the address compared to what is listed in MAXIS."
				Text 185, 75, 20, 10, "<<---"
				Text 200, 75, 200, 25, "This is the current residence address known in MAXIS"
				If mail_line_one = "" Then
					Text 185, 195, 20, 10, "<<---"
					Text 200, 195, 200, 25, "MAXIS does not have a mailing address listed."
				Else
					Text 185, 210, 20, 10, "<<---"
					Text 200, 210, 200, 25, "This is the current mailing addresss known in MAXIS"
				End If
			ElseIf dlg_step = "STEP 2" Then
				GroupBox 180, 5, 300, 260, "Client Conversation"
				Text 185, 20, 245, 10, "^^2 - Read the Residence Address to the client."
				Text 195, 30, 125, 10, "Ask: Is thatthe address you live at?"
				GroupBox 195, 36, 280, 20,""
				Text 200, 43, 100, 20,"Select or Type client response"
				GroupBox 345, 52, 130, 20,""
				Text 360, 58, 100, 10, "Record specific changes here"

				Text 165, 20, 20, 10, "--->>"
				Text 15, 20, 150, 25, "Now that the addresses on the CAF and in MAXIS have been reconciled, have the client verbally confirm the address."
			End If
		End If

		PushButton 465, 365, 80, 15, "Back to Update", return_to_full_dialog_btn
	EndDialog

	dialog Dialog1
end function

function define_main_dialog()

' BeginDialog Dialog1, 0, 0, 550, 385, "Full Interview Questions"
'   ButtonGroup ButtonPressed
'     PushButton 415, 365, 50, 15, "NEXT", next_btn
'     PushButton 465, 365, 80, 15, "Complete Interview", finish_interview_btn
'     PushButton 485, 10, 60, 15, "CAF Page 1", caf_page_one_btn
'     PushButton 485, 135, 60, 15, "CAF Page 1", Button27
'   Text 495, 265, 60, 15, "Q. 20 and 21"
' EndDialog

	BeginDialog Dialog1, 0, 0, 550, 385, "Full Interview Questions"

	  ButtonGroup ButtonPressed
		If page_display = show_pg_one Then										'THIS IS THE INFORMATION FOR PAGE  ----------------------------------------------------------------------------------------------------------------------------------
			Text 495, 12, 60, 13, "CAF Page 1"

			GroupBox 180, 5, 300, 260, "Client Conversation"
			PushButton 182, 20, 18, 10, "^^2", page_1_step_2_btn
			Text 200, 20, 245, 10, " - Read the Residence Address to the client."
			Text 195, 30, 125, 10, "Ask: Is that the address you live at?"
			ComboBox 195, 40, 280, 45, "Select or Type client response"+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Household is Homeless", residence_address_response
			PushButton 345, 55, 130, 13, "Record specific changes here", client_verbal_changes_resi_address_btn
			' Text 185, 70, 245, 10, "^^3 - Ask: Are you experiencing housing instability (or homelessness)?"
			PushButton 182, 70, 18, 10, "^^3", page_1_step_3_btn
			Text 200, 70, 245, 10, " - Ask: Are you experiencing housing instability (or homelessness)?"
			ComboBox 195, 80, 280, 45, "Select or Type client response"+chr(9)+"Yes"+chr(9)+"No", household_homeless_response
			' Text 185, 100, 50, 10, "^^4 - Explain:"
			PushButton 182, 100, 18, 10, "^^4", page_1_step_4_btn
			Text 200, 100, 50, 10, " - Explain:"
			Text 200, 110, 275, 25, "We use mail as our primary means of communication to let you know if any action is required for any benefits you receive to continue. It is important the address we send mail to has your name listed and that you check it regularly. "
			Text 200, 140, 240, 15, "If a mailing address has been listed on the CAF or in MAXIS - read it to the client.                                         Which Address?"
			DropListBox 365, 150, 110, 45, "Residence Address"+chr(9)+"Mailing Address", which_address_are_we_discussing
			Text 195, 165, 150, 10, "Ask: Can you receive mail at this address?"
			ComboBox 195, 175, 280, 45, "Select or Type client response"+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"No - Use General Delivery", mail_received_at_this_address_response
			PushButton 345, 190, 130, 13, "Record specific changes here", client_verbal_changes_mail_address_btn
			' Text 185, 205, 140, 10, "^^5 - If General Delivery requested - Explain:"
			PushButton 182, 205, 18, 10, "^^5", page_1_step_5_btn
			Text 200, 205, 140, 10, " - If General Delivery requested - Explain:"
			Text 195, 215, 260, 25, "GD can be used to have your mail held at the post office. You will need a photo ID to collect mail from GD. As our mail often requires a response within 10 days, you should be checking GD at least every 2-3 days."
			DropListBox 275, 245, 200, 45, "General Delivery Not Requested"+chr(9)+"Explained and Client Confirmed Understanding", confirm_gen_del_explanation
			GroupBox 10, 5, 170, 260, "Address Information listed in MAXIS"
			' Text 15, 20, 160, 25, "^^1 - Compare this address to the one entered on the CAF. If there is a difference, press the button below to update."
			Text 15, 20, 160, 25, "        - Compare this address to the one entered on the CAF. If there is a difference, press the button below to update."
			PushButton 12, 18, 18, 10, "^^1", page_1_step_1_btn
			Text 20, 60, 70, 10, "Residence Address"
			Text 25, 75, 150, 10, resi_line_one
			If resi_line_two = "" Then
				Text 25, 85, 150, 10, resi_city & ", " & resi_state & " " & resi_zip
				Text 25, 95, 150, 10, "County: " & resi_county

			Else
				Text 25, 85, 150, 10, resi_line_two
				Text 25, 95, 150, 10, resi_city & ", " & resi_state & " " & resi_zip
				Text 25, 105, 150, 10, "County: " & resi_county
			End If
			Text 25, 120, 65, 10, "Homeless: " & homeless
			Text 25, 130, 150, 20, "Living Situation: " & living_sit
			Text 25, 155, 60, 10, "IND RES - " & ind_reservation
			Text 25, 165, 150, 10, "RES NAME - " & res_name
			Text 25, 180, 150, 10, "Verification: " & verif
			If mail_line_one = "" Then
				Text 20, 195, 150, 10, "No MAILING ADDRESS Listed"
			ElseIf mail_line_two = "" Then
				Text 20, 195, 70, 10, "Mailing Address"
				Text 25, 210, 150, 10, mail_line_one
				Text 25, 220, 150, 10, mail_city & ", " & mail_state & " " & mail_zip
			Else
				Text 20, 195, 70, 10, "Mailing Address"
				Text 25, 210, 150, 10, mail_line_one
				Text 25, 220, 150, 10, mail_line_two
				Text 25, 230, 150, 10, mail_city & ", " & mail_state & " " & mail_zip
			End If
			PushButton 50, 245, 125, 15, "CAF has Different Information", caf_info_different_btn
			GroupBox 10, 265, 470, 80, "Phone Contact"
			Text 20, 280, 115, 10, "Current Phone Numbers in MAXIS"
			Text 25, 295, 60, 10, phone_numb_one
			Text 90, 295, 35, 10, phone_type_one
			Text 25, 310, 60, 10, phone_numb_two
			Text 90, 310, 35, 10, phone_type_two
			Text 25, 325, 60, 10, phone_numb_three
			Text 90, 325, 40, 10, phone_type_three
			' Text 150, 280, 185, 10, "^^6 - Ask: What is the best phone number to reach out at?"
			PushButton 147, 280, 18, 10, "^^6", page_1_step_6_btn
			Text 165, 280, 185, 10, " - Ask: What is the best phone number to reach out at?"
			EditBox 340, 275, 65, 15, reported_phone_number
			Text 185, 295, 120, 10, "What type of phone number is this?"
			DropListBox 305, 290, 100, 45, "Select One..."+chr(9)+"Cell"+chr(9)+"Home"+chr(9)+"Work"+chr(9)+"Message Only"+chr(9)+"TTY/TDD", reported_phone_type
			' Text 150, 315, 320, 10, "^^7 - For each number on the lest, Read it to the client and Ask: is this still agood number?"
			PushButton 147, 315, 18, 10, "^^7", page_1_step_7_btn
			Text 165, 315, 320, 10, " - For each number on the lest, Read it to the client and Ask: is this still agood number?"
			PushButton 345, 325, 130, 13, "Record Any Changes to Numbers Here", client_verbal_changes_phone_numbers_btn

			GroupBox 10, 345, 350, 30, "Living Situation"
			' Text 20, 360, 125, 10, "^^8 - Ask: What is your living situation?"
			PushButton 17, 360, 18, 10, "^^8", page_1_step_8_btn
			Text 35, 360, 125, 10, " - Ask: What is your living situation?"
			DropListBox 150, 355, 200, 15, "Select One..."+chr(9)+"Own Housing(lease, mortgage, or roomate)"+chr(9)+"Family/Friends due to economic hardship"+chr(9)+"Servc prvdr- foster/group home"+chr(9)+"Hospital/Treatment/Detox/Nursing Home"+chr(9)+"Jail/Prison//Juvenile Det."+chr(9)+"Hotel/Motel"+chr(9)+"Emergency Shelter"+chr(9)+"Place not meant for Housing"+chr(9)+"Declined"+chr(9)+"Unknown", clt_response_living_sit
		End If
		If page_display = show_pg_memb_list Then										'THIS IS THE INFORMATION FOR PAGE  ----------------------------------------------------------------------------------------------------------------------------------
			Text 495, 27, 60, 13, "CAF MEMBs"

			' Text 15, 15, 450, 10, "THE QUESTION GOES HERE"
			' Text 135, 35, 65, 10, "Answer on the CAF"
			' Text 265, 35, 70, 10, "Confirm CAF Answer"
			' DropListBox 200, 30, 40, 45, "No"+chr(9)+"Yes", caf_answer
			' ComboBox 340, 30, 125, 45, "", confirm_caf_answer
			grp_box_len = 100 + UBound(HH_MEMB_ARRAY, 1) * 15

			GroupBox 10, 30, 470, grp_box_len, "Household Members Listed in MAXIS"
			GroupBox 410, 70, 60, grp_box_len - 45, "REMO Date"
			Text 15, 45, 240, 10, "^^1 - Ask: Please list everyone that lives at the address/lives with you."
			Text 30, 55, 250, 10, "Check the boxes: Is this member listed on the CAF? and Reported Verbally?"
			Text 20, 75, 35, 10, "Applicant:"
			Text 25, 90, 255, 10, "- MEMB " & HH_MEMB_ARRAY(0).ref_number & "    " & HH_MEMB_ARRAY(0).full_name
			CheckBox 285, 90, 50, 10, "On the CAF", HH_MEMB_ARRAY(0).checkbox_one
			CheckBox 340, 90, 70, 10, "Verbally Reported", HH_MEMB_ARRAY(0).checkbox_two
			EditBox 415, 85, 50, 15, HH_MEMB_ARRAY(0).left_hh_date

			Text 20, 110, 95, 10, "Other Household Members:"
			y_pos = 125
			for i = 1 to UBound(HH_MEMB_ARRAY, 1)
				Text 25, y_pos, 255, 10, "- MEMB " & HH_MEMB_ARRAY(i).ref_number & "    " & HH_MEMB_ARRAY(i).full_name & " - " & HH_MEMB_ARRAY(i).rel_to_applcnt & " of Memb 01"
				CheckBox 285, y_pos, 50, 10, "On the CAF", HH_MEMB_ARRAY(i).checkbox_one
				CheckBox 340, y_pos, 70, 10, "Verbally Reported", HH_MEMB_ARRAY(i).checkbox_two
				EditBox 415, y_pos - 5, 50, 15, HH_MEMB_ARRAY(i).left_hh_date
				y_pos = y_pos + 15
			Next
			PushButton 360, 40, 115, 15, "Add Member not Listed Here", add_memb_to_list_btn
			y_pos = y_pos + 10
			Text 15, y_pos, 345, 10, "^^2 - If any members were not reported, Ask: About each member and if they have left the household."
			y_pos = y_pos + 20
			Text 15, y_pos, 465, 10, "Current AREP in MAXIS: FIRST NAME LAST NAME ALL THE THINGS "
			y_pos = y_pos + 15
			Text 15, y_pos, 185, 10, "^^3 - Check the CAF to see if an Authorized Rep is listed. "
			ComboBox 210, y_pos - 5, 270, 45, " "+chr(9)+"Yes - I would like an AREP"+chr(9)+"No - Do not authorize someone on my case", arep_response
			y_pos = y_pos + 10
			Text 30, y_pos, 245, 10, "Ask: Do you want another person authorized to talk to us about your case?"
			y_pos = y_pos + 20
			Text 15, y_pos, 465, 10, "Current SWKR in MAXIS: FIRST NAME LAST NAME ALL THE THINGS "
			y_pos = y_pos + 15
			Text 15, y_pos, 175, 10, "^^4 - Check the CAF to see if a Social Worker is listed. "
			ComboBox 210, y_pos - 5, 270, 40, " "+chr(9)+"Yes - I would like an AREP"+chr(9)+"No - Do not authorize someone on my case", Combo2
			y_pos = y_pos + 10
			Text 30, y_pos, 275, 10, "Ask: Do you have a social worker you want authorized to talk to us about your case?"
			' PushButton 10, 10, 80, 15, "List of HH Members", hh_list_btn
			Text 20, 8, 80, 15, "List of HH Members"
			PushButton 95, 5, 125, 15, "Review HH Member Information", HH_memb_detail_review
			If case_has_imig = TRUE Then PushButton 225, 5, 125, 15, "Immigration Information", hh_imig_btn
			If case_has_imig = FALSE Then Text 230, 8, 125, 10, "All MEMBERS listed as Citizens"
		End If
		If page_display = show_pg_memb_info Then										'THIS IS THE INFORMATION FOR PAGE  ----------------------------------------------------------------------------------------------------------------------------------
			Text 495, 27, 60, 13, "CAF MEMBs"

			Text 15, 23, 460, 10, "^^1 - Review the personal information/detail for each household member on this case with the client. Review and add ALL household relationships."
			Text 20, 33, 460, 10, "* Be sure to check if proof of identity is required and look in ECF or SOL-Q to ensure verification is correct."
			Text 20, 43, 460, 10, "* Confirm name spelling, language, marital status, immigration/citizenship status."
			Text 20, 53, 460, 10, "* If the SSN has not been validated, ask the client for the correct SSN."
			grp_len = 195
			For the_rela = 0 to UBound(ALL_HH_RELATIONSHIPS_ARRAY, 2)
				If ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_ref, the_rela) = HH_MEMB_ARRAY(memb_selected).ref_number Then grp_len = grp_len + 20
			Next
			If grp_len = 195 Then grp_len = 215
			GroupBox 50, 65, 425, grp_len, "Information for " & HH_MEMB_ARRAY(memb_selected).full_name
			Text 60, 80, 165, 10, "Name: " & HH_MEMB_ARRAY(memb_selected).full_name
			Text 225, 80, 85, 10, "Review of info response:"
			ComboBox 315, 75, 155, 45, "", review_memb_info_detail
			Text 75, 90, 165, 10, "Age: " & HH_MEMB_ARRAY(memb_selected).age & "      DOB: "  & HH_MEMB_ARRAY(memb_selected).date_of_birth
			Text 75, 100, 235, 10, "Written Language: " & HH_MEMB_ARRAY(memb_selected).written_lang & "       Spoken Language: " & HH_MEMB_ARRAY(memb_selected).spoken_lang
			Text 75, 115, 185, 10, "Proof of Identity:" & HH_MEMB_ARRAY(memb_selected).id_verif
			Text 275, 120, 100, 10, "Identity Proof:"
			ComboBox 330, 115, 135, 45, "Select or Type"+chr(9)+"Not Needed"+chr(9)+"Requested"+chr(9)+"ECF"+chr(9)+"SMI/SOL-Q", identity_proof_found
			CheckBox 80, 125, 140, 10, "Check Here if identity proof is required", id_required_checkbox
			Text 75, 140, 65, 10, "SSN: " & HH_MEMB_ARRAY(memb_selected).ssn
			If left(HH_MEMB_ARRAY(memb_selected).ssn_verif, 1) <> "V" Then Text 155, 140, 195, 10, "SSN has not been validated, review this number with client:"
			If left(HH_MEMB_ARRAY(memb_selected).ssn_verif, 1) <> "V" Then ComboBox 350, 135, 115, 45, "SSN wrong - now updated"+chr(9)+"This number is correct", non_validated_ssn_detail
			If HH_MEMB_ARRAY(memb_selected).spouse_ref <> "" Then Text 60, 155, 400, 10, "Marital Status: " & HH_MEMB_ARRAY(memb_selected).marital_status & "          Spouse's Name: " & HH_MEMB_ARRAY(memb_selected).spouse_ref & " - " & HH_MEMB_ARRAY(memb_selected).spouse_name
			If HH_MEMB_ARRAY(memb_selected).spouse_ref = "" Then Text 60, 155, 350, 10, "Marital Status: " & HH_MEMB_ARRAY(memb_selected).marital_status
			Text 60, 170, 140, 10, "In Minnesota for at least 12 Months: " & HH_MEMB_ARRAY(memb_selected).in_mn_12_mo
			If HH_MEMB_ARRAY(memb_selected).in_mn_12_mo = "No" Then Text 215, 170, 160, 10, "Arrived from: " & HH_MEMB_ARRAY(memb_selected).former_state & "     Arrival Date: " & HH_MEMB_ARRAY(memb_selected).mn_entry_date

			Text 60, 185, 65, 10, "US Citizen: " & HH_MEMB_ARRAY(memb_selected).citizen
			If HH_MEMB_ARRAY(memb_selected).imig_exists = TRUE THen
				Text 65, 200, 410, 10, "Immigration Status: " & HH_MEMB_ARRAY(memb_selected).imig_status & "     Immigration Documentaion: " & HH_MEMB_ARRAY(memb_selected).imig_status_verif
				' Text 200, 200, 250, 10, "Immigration Documentaion: " & HH_MEMB_ARRAY(memb_selected).imig_status_verif
			End If
			Text 60, 215, 175, 10, "Relationships in the Household: This Person is ..."
			Text 295, 215, 160, 10, "Relationship Verification:"
			y_pos = 230
			For the_rela = 0 to UBound(ALL_HH_RELATIONSHIPS_ARRAY, 2)
				If ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_ref, the_rela) = HH_MEMB_ARRAY(memb_selected).ref_number Then
					DropListBox 75, y_pos, 80, 45,rela_type_dropdown, ALL_HH_RELATIONSHIPS_ARRAY(rela_type, the_rela)
					Text 165, y_pos + 5, 10, 10, "to"
					DropListBox 185, y_pos, 100, 45, "Select One..."&memb_droplist, ALL_HH_RELATIONSHIPS_ARRAY(rela_pers_two, the_rela)
					CheckBox 295, y_pos+ 5, 40, 10, "Required", ALL_HH_RELATIONSHIPS_ARRAY(verif_req_checkbox, the_rela)
					ComboBox 345, y_pos, 110, 45, rela_verif_dropdown & ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, the_rela), ALL_HH_RELATIONSHIPS_ARRAY(rela_verif, the_rela)
					y_pos = y_pos + 20
				End If
			Next
			If y_pos = 230 Then
				Text 75, y_pos, 300, 10, "No relationships between household members known in MAXIS at this time."
				y_pos = y_pos + 20
			End If
			' DropListBox 75, 230, 80, 45, ""+chr(9)+"Parent"+chr(9)+"Child"+chr(9)+"Sibling"+chr(9)+"Spouse"+chr(9)+"Grandparent"+chr(9)+"Neice"+chr(9)+"Nephew"+chr(9)+"Aunt"+chr(9)+"Uncle"+chr(9)+"Grandchild", relationship_type
			' Text 165, 235, 10, 10, "to"
			' DropListBox 185, 230, 100, 45, "", relationship_member
			' CheckBox 295, 230, 40, 10, "Required", verif_req_checkbox
			' ComboBox 345, 230, 110, 45, "", relationship_verif_type
			' DropListBox 75, 250, 80, 45, "", List3
			' Text 165, 250, 10, 10, "to"
			' DropListBox 185, 250, 100, 45, "", List4
			' CheckBox 295, 250, 40, 10, "Required", Check3
			' ComboBox 345, 250, 110, 45, "", Combo5

			PushButton 75, y_pos, 90, 10, "Add Another Relationship", add_relationship_btn
			y_pos = y_pos + 15
			PushButton 55, y_pos, 125, 10, "Update Member Information", memb_info_change
			PushButton 410, y_pos, 60, 10, "NEXT MEMB", next_memb_btn

			Text 50, 370, 75, 10, "Principal Wage Earner: "
			DropListBox 130, 365, 125, 15, memb_droplist, the_pwe_for_this_case

			btn_pos = 70
			For i = 0 to Ubound(HH_MEMB_ARRAY, 1)
				If i = memb_selected Then
					Text 9, btn_pos+1, 40, 10, "MEMB " & HH_MEMB_ARRAY(i).ref_number
				Else
					PushButton 5, btn_pos, 40, 10, "MEMB " & HH_MEMB_ARRAY(i).ref_number, HH_MEMB_ARRAY(i).button_one
				End If
				btn_pos = btn_pos + 10
			Next
			' PushButton 5, 70, 40, 10, "MEMB 01", memb_select_btn
			' PushButton 5, 80, 40, 10, "MEMB 01", Button9
			' PushButton 5, 90, 40, 10, "MEMB 01", Button10
			' PushButton 5, 100, 40, 10, "MEMB 01", Button11
			' PushButton 5, 110, 40, 10, "MEMB 01", Button12

			' PushButton 95, 10, 125, 15, "Review HH Member Information", HH_memb_detail_review
			PushButton 10, 5, 80, 15, "List of HH Members", hh_list_btn
			Text 105, 8, 125, 15, "Review HH Member Information"
			If case_has_imig = TRUE Then PushButton 225, 5, 125, 15, "Immigration Information", hh_imig_btn
			If case_has_imig = FALSE Then Text 230, 8, 125, 10, "All MEMBERS listed as Citizens"
		End If
		If page_display = show_pg_imig Then										'THIS IS THE INFORMATION FOR PAGE  ----------------------------------------------------------------------------------------------------------------------------------
			Text 495, 27, 60, 13, "CAF MEMBs"

			Text 15, 30, 460, 10, "DISCUSS Immigration details for each household member."
			' Text 20, 33, 460, 10, "* Be sure to check if proof of identity is required and look in ECF or SOL-Q to ensure verification is correct."
			' Text 20, 43, 460, 10, "* Confirm name spelling, language, marital status, immigration/citizenship status."
			' Text 20, 53, 460, 10, "* If the SSN has not been validated, ask the client for the correct SSN."

			GroupBox 50, 45, 425, 115, "Immmigration Information for " & HH_MEMB_ARRAY(memb_selected).full_name
			If HH_MEMB_ARRAY(memb_selected).imig_exists = TRUE Then
				' Text 85, 80, 165, 10, "Name: " & HH_MEMB_ARRAY(memb_selected).full_name
				Text 230, 45, 40, 10, "Alien ID:"
				Text 270, 45, 50, 10, HH_MEMB_ARRAY(memb_selected).alien_id_nbr

				Text 85, 60, 40, 10, "IMIG Status:"
				Text 130, 60, 150, 10, HH_MEMB_ARRAY(memb_selected).imig_status
				' DropListBox 125, 95, 200, 45, "Select One..."+chr(9)+imig_status_dropdown, HH_MEMB_ARRAY(memb_selected).imig_status
				Text 230, 60, 40, 10, "Status Date:"
				Text 270, 60, 40, 10, HH_MEMB_ARRAY(memb_selected).imig_status_date
				' EditBox 380, 95, 75, 15, HH_MEMB_ARRAY(memb_selected).imig_status_date
				If HH_MEMB_ARRAY(memb_selected).imig_status = "24 - LPR" Then
					Text 330, 60, 60, 10, "LPR adjusted from:"
					Text 395, 75, 75, 10, HH_MEMB_ARRAY(memb_selected).lpr_adj_from
				End If

				Text 85, 75, 40, 10, "Status Verif:"
				Text 130, 75, 200, 10, HH_MEMB_ARRAY(memb_selected).imig_status_verif
				' EditBox 125, 115, 200, 15, HH_MEMB_ARRAY(memb_selected).imig_status_verif
				Text 230, 75, 40, 10, "Document:"
				Text 270, 75, 200, 10, HH_MEMB_ARRAY(memb_selected).imig_active_doc
				Text 270, 75, 200, 10, HH_MEMB_ARRAY(memb_selected).imig_recvd_doc

				Text 70, 90, 65, 10, "Date of US Entry:"
				Text 130, 90, 75, 10, HH_MEMB_ARRAY(memb_selected).us_entry_date
				' EditBox 140, 135, 75, 15, HH_MEMB_ARRAY(memb_selected).us_entry_date

				Text 230, 90, 35, 10, "Nationality:"
				Text 270, 90, 75, 10, HH_MEMB_ARRAY(memb_selected).nationality
				' DropListBox 140, 155, 75, 15, "Select One..."+chr(9)+imig_nationality_droplist, HH_MEMB_ARRAY(memb_selected).nationality
				If left(HH_MEMB_ARRAY(memb_selected).nationality, 2) = "OT" Then
					Text 395, 90, 50, 10, "Detail, if other:"
					Text 465, 90, 75, 10, HH_MEMB_ARRAY(memb_selected).nationality_detail
					' DropListBox 140, 155, 75, 15, "Select One..."+chr(9)+imig_nationality_droplist, HH_MEMB_ARRAY(memb_selected).nationality
				End If
				'WE CAN LOOK UP A I-94 RECORD OURSELVES!!!'
				' Text 60, 180, 350, 10, "^^2 - ASK "

				If HH_MEMB_ARRAY(memb_selected).spon_exists = TRUE Then
					Text 85, 105, 80, 10, "Client has sponsor."
					Text 230, 105, 200, 10, HH_MEMB_ARRAY(memb_selected).spon_name
				Else
					Text 85, 105, 80, 10, "No sponsor."
				End If

				Text 50, 165, 400, 10, "^^1 - ASK all questions in this series:"
				GroupBox 50, 180, 425, 175, "Immigration Questions - Step " & imig_questions_step

				If imig_questions_step = 1 Then

					Text 70, 200, 150, 10, "What is your date of entry to the United States?"
					EditBox 225, 195, 75, 15, HH_MEMB_ARRAY(memb_selected).us_entry_date

					Text 70, 220, 85, 10, "What is your nationality?"
					DropListBox 155, 215, 125, 15, "Select One..."+chr(9)+imig_nationality_droplist, HH_MEMB_ARRAY(memb_selected).nationality
					Text 285, 220, 75, 10, "If 'Other' add details"
					EditBox 360, 215, 100, 15, HH_MEMB_ARRAY(memb_selected).nationality_detail

					Text 70, 240, 150, 10, "What is your current immigration status?"
					ComboBox 225, 235, 150, 15, "Type or Select"+chr(9)+imig_status_dropdown, HH_MEMB_ARRAY(memb_selected).imig_status

				ElseIf imig_questions_step = 2 Then
					Text 70, 200, 300, 10, "Since this client is an LPR, we need to clarify details of the status."

					Text 70, 220, 150, 10, "What was your previous immigration status?"
					ComboBox 225, 215, 150, 15, "Type or Select"+chr(9)+imig_status_dropdown, HH_MEMB_ARRAY(memb_selected).imig_prev_status

					Text 70, 240, 150, 10, "What date did you adjust to LPR?"
					EditBox 225, 235, 75, 15, HH_MEMB_ARRAY(memb_selected).us_entry_dat
				ElseIf imig_questions_step = 3 Then
					Text 70, 200, 200, 10, "What is your Immigration documentation do you currently have"
					ComboBox 275, 195, 150, 15, "Type or Select"+chr(9)+imig_document_dropdown, HH_MEMB_ARRAY(memb_selected).imig_clt_current_doc

					Text 70, 220, 200, 10, "What Immigration documentation do we have on file?"
					ComboBox 275, 215, 150, 15, "Type or Select"+chr(9)+imig_document_dropdown, HH_MEMB_ARRAY(memb_selected).imig_doc_on_file

					Text 70, 240, 100, 10, "Has SAVE been completed?"
					DropListBox 175, 235, 75, 15, "No"+chr(9)+"Yes", HH_MEMB_ARRAY(memb_selected).imig_save_completed

					Text 70, 240, 75, 10, "Do you have a sponsor?"
					DropListBox 150, 235, 150, 15, "No"+chr(9)+"Yes", HH_MEMB_ARRAY(memb_selected).ask_about_spon
				ElseIf imig_questions_step = 4 Then
				ElseIf imig_questions_step = 5 Then
				End If
				If imig_questions_step <> 1 Then PushButton 50, 342, 75, 13, "PREVIOUS", prev_imig_btn
				PushButton 400, 342, 75, 13, "NEXT", next_imig_btn

				' Immigration Questions
				' STEP 1
				' 	What is your date of entry?
				' 	What is your nationality? If other, fill edit box.
				' 	What is your current immigration status?
				' STEP 2
				' 	IF LPR, what status did you have previously?
				' 	IF LPR, when did you adjust to LPR?
				' STEP 3
				' 	What immigration document do you have to support your status?
				' 	What immigration document do we have on file?
				' 	Has a SAVE been completed?
				' 	IF LPR and not adjusted from refugee/asylee, Do you have a sponsor?
				' STEP 4
				' 	IF Sponsor:
				' 	Sponsor name?
				' 	Sponsor’s address?
				' 	Is sponsor married?
				' 	Does the sponsor have any children that live in their household?
				' STEP 5
				' 	What is the sponsor’s income?
				' 	What is the sponsor’s spouse’s income?


			Else
				Text 60, 60, 165, 10, "Name: " & HH_MEMB_ARRAY(memb_selected).full_name
				Text 60, 80, 200, 10, "There is no IMMIGRATION information for " & HH_MEMB_ARRAY(memb_selected).full_name
			End If
			' PushButton 400, 55, 70, 13, "UPDATE DETAIL", update_imig_info_btn

			btn_pos = 50
			For i = 0 to Ubound(HH_MEMB_ARRAY, 1)
				If i = memb_selected Then
					Text 9, btn_pos+1, 40, 10, "MEMB " & HH_MEMB_ARRAY(i).ref_number
				Else
					PushButton 5, btn_pos, 40, 10, "MEMB " & HH_MEMB_ARRAY(i).ref_number, HH_MEMB_ARRAY(i).button_one
				End If
				btn_pos = btn_pos + 10
			Next
			PushButton 5, btn_pos, 40, 20, "NEXT MEMB", next_memb_btn
			PushButton 10, 5, 80, 15, "List of HH Members", hh_list_btn
			PushButton 95, 5, 125, 15, "Review HH Member Information", HH_memb_detail_review
			Text 240, 8, 125, 10, "Immigration Information"
		End If
		If page_display = show_q_1_2 Then										'THIS IS THE INFORMATION FOR PAGE  ----------------------------------------------------------------------------------------------------------------------------------
			Text 500, 42, 60, 13, "Q. 1 and 2"

			Text 5, 10, 330, 10, "^^1 - Enter the answers listed on the actual CAF form for Q1 and Q2 into the 'Answer on the CAF' field."
		    Text 5, 25, 370, 10, "^^2 - ASK - Q1 and Q2 and record the verbal answers in the 'Confirm CAF Answer' field under the question."
		    Text 15, 40, 235, 10, "Q. 1. Does everyone in your household buy, fix or eat food with you?"
		    Text 370, 40, 65, 10, "Answer on the CAF"
		    DropListBox 435, 35, 40, 45, caf_answer_droplist, q1_caf_answer
		    Text 35, 60, 70, 10, "Confirm CAF Answer"
		    ComboBox 110, 55, 365, 45, "", q1_confirm_caf_answer
		    Text 15, 85, 315, 20, "Q. 2. Is anyone who is in the household, who is age 60 or over or disabled, unable to buy or fix food due to a disability?"
		    Text 370, 85, 65, 10, "Answer on the CAF"
		    DropListBox 435, 80, 40, 45, caf_answer_droplist, q2_caf_answer
		    Text 35, 110, 70, 10, "Confirm CAF Answer"
		    ComboBox 110, 105, 365, 45, "", q2_confirm_caf_answer
		    Text 5, 140, 285, 10, "^^3 - ASK - Is there anyone else living in the house that does NOT share food with you?"
		    Text 20, 155, 105, 10, "Anyone else NOT sharing food?"
		    ComboBox 130, 150, 345, 45, "", anyone_else_in_hh_confirm
		    Text 5, 190, 255, 10, "^^4 - Using the above questions, CONFIRM the information below from MAXIS"
		    Text 20, 205, 455, 10, "HH Members UNABLE to P and P Seperately: " & members_unable_to_fix_food

			y_pos = 235
			grp_len = 45
			If group_one_number = "__" Then
		    	Text 30, y_pos, 440, 10, "** No seperate Groups - everyone purchases and prepares together."
				y_pos = y_pos + 15
			Else
				Text 30, y_pos, 440, 10, "Group: " & group_one_number & " - " & group_one_member_list
				y_pos = y_pos + 15
				If group_two_number <> "__" Then
					Text 30, y_pos, 440, 10, "Group: " & group_two_number & " - " & group_two_member_list
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If
				If group_three_number <> "__" Then
					Text 30, y_pos, 440, 10, "Group: " & group_three_number & " - " & group_three_member_list
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If
				If group_four_number <> "__" Then
					Text 30, y_pos, 440, 10, "Group: " & group_four_number & " - " & group_four_member_list
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If
				If group_five_number <> "__" Then
					Text 30, y_pos, 440, 10, "Group: " & group_five_number & " - " & group_five_member_list
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If
			End If

			GroupBox 20, 220, 455, grp_len, "Eating Groups (who purchases and prepares together on this case)"
		    PushButton 405, y_pos, 65, 10, "Update Groups", update_groups_btn
			y_pos = y_pos + 35
		    Text 5, y_pos, 215, 10, "^^5 - Confirm information for Q1 and Q2 are complete and correct:"
		    ComboBox 25, y_pos + 10, 450, 45, "Type or Select", q1_and_q2_confirmation


		End If
		If page_display = show_q_3 Then										'THIS IS THE INFORMATION FOR PAGE  ----------------------------------------------------------------------------------------------------------------------------------
			Text 508, 57, 60, 13, "Q. 3"

			Text 5, 10, 305, 10, "^^1 - Enter the answers listed on the actual CAF form for Q3 into the 'Answer on the CAF' field."
			Text 15, 25, 180, 10, "Q. 3. Is anyone in your household attending school?"
			Text 370, 25, 65, 10, "Answer on the CAF"
			DropListBox 435, 20, 40, 45, caf_answer_droplist, q3_caf_answer
			Text 5, 45, 405, 10, "^^2 - ASK - 'Is anyone attending school?' and record the verbal answers in the 'Confirm CAF Answer' field under the question."
			Text 40, 65, 70, 10, "Confirm CAF Answer"
			ComboBox 110, 60, 365, 45, "", q3_confirm_caf_answer


		    Text 5, 85, 380, 10, "^^3 - If there are school-age children in the household - ASK - What grade and school district does each child attend?"
		    Text 20, 95, 460, 10, "Child (Name and Age) --------------------------------------------------Grade ---------------------District -------Verification -----------------------------------------Status ---------------------"

			y_pos = 105
			for i = 0 to UBound(HH_MEMB_ARRAY, 1)
				If HH_MEMB_ARRAY(i).age < 19 AND HH_MEMB_ARRAY(i).age > 4 Then
					Text 20, y_pos + 5, 170, 10, "MEMB " & HH_MEMB_ARRAY(i).ref_number & " - " & HH_MEMB_ARRAY(i).full_name & " - Age: " & HH_MEMB_ARRAY(i).age
					DropListBox 190, y_pos, 60, 45, grade_droplist, HH_MEMB_ARRAY(i).school_grade
					EditBox 255, y_pos, 35, 15, HH_MEMB_ARRAY(i).school_district
					ComboBox 295, y_pos, 115, 45, schl_ver_droplist & chr(9)+HH_MEMB_ARRAY(i).school_verif, HH_MEMB_ARRAY(i).school_verif
					DropListBox 415, y_pos, 60, 45, schl_status_droplist, HH_MEMB_ARRAY(i).school_status
					y_pos = y_pos + 20
				End If
			Next
			If y_pos = 105 Then
				Text 20, y_pos + 5, 200, 10, "NO Children age 5 - 18 known on this case."
				y_pos = y_pos + 20
			End If
		    PushButton 385, 82, 95, 13, "Search School Districts", search_district_btn

			y_pos = y_pos + 10
			Text 5, y_pos, 230, 10, "^^4 - ASK - Is anyone attending college/university or other higher ed?"
			y_pos = y_pos + 10
		    ComboBox 20, y_pos, 455, 45, "", school_higher_ed_answer
			y_pos = y_pos + 20
		    Text 5, y_pos, 320, 10, "^^5 - If anyone is attenting hight ed, ENTER or CONFIRM information for the household members:"
			hi_ed_btn_pos = y_pos - 3
			y_pos = y_pos + 10
		    Text 20, y_pos, 460, 10, "Household Member ------------------------------------------------School ---------------------------------------------Status ---------------------Verification --------------------------------------"
			y_pos = y_pos + 10
			start_pos = y_pos

			for i = 0 to UBound(HH_MEMB_ARRAY, 1)
				If left(HH_MEMB_ARRAY(i).school_type, 2) = "08" OR left(HH_MEMB_ARRAY(i).school_type, 2) = "09" OR left(HH_MEMB_ARRAY(i).school_type, 2) = "10" Then
				    Text 20, y_pos + 5, 160, 10, "MEMB " & HH_MEMB_ARRAY(i).ref_number & " - " & HH_MEMB_ARRAY(i).full_name & " - Age: " & HH_MEMB_ARRAY(i).age
				    EditBox 180, y_pos, 105, 15, HH_MEMB_ARRAY(i).school_name
				    DropListBox 295, y_pos, 60, 45, schl_status_droplist, HH_MEMB_ARRAY(i).school_status
				    ComboBox 360, y_pos, 115, 45, schl_ver_droplist & chr(9)+HH_MEMB_ARRAY(i).school_verif, HH_MEMB_ARRAY(i).school_verif
					y_pos = y_pos + 20
				End If
			Next
			If y_pos = start_pos Then
				Text 20, y_pos + 5, 200, 10, "No Post Secondary students known on this case."
				y_pos = y_pos + 20
			End if
			PushButton 355, hi_ed_btn_pos, 120, 13, "Add Higher Ed Student", add_higher_ed_studen

			y_pos = y_pos + 5
		    Text 5, y_pos, 235, 10, "^^6 - ASK - Is anyone attending GED/ELL (English Language Learning)?"
			y_pos = y_pos + 10
		    ComboBox 20, y_pos, 455, 45, "", Combo5
			y_pos = y_pos + 20
		    Text 5, y_pos, 305, 10, "^^7 - If anyone is in GED or ELL, ENTER or CONFRIM information for the household members:"
			ged_ell_btn_pos = y_pos - 3
			y_pos = y_pos + 10
		    Text 20, y_pos, 460, 10, "Household Member ------------------------------------------------School ---------------------------------------------Status ---------------------Verification --------------------------------------"
			y_pos = y_pos + 10
			start_pos = y_pos

			for i = 0 to UBound(HH_MEMB_ARRAY, 1)
				If left(HH_MEMB_ARRAY(i).school_type, 2) = "03" OR left(HH_MEMB_ARRAY(i).school_type, 2) = "13" Then
				    Text 20, y_pos + 5, 160, 10, "MEMB " & HH_MEMB_ARRAY(i).ref_number & " - " & HH_MEMB_ARRAY(i).full_name & " - Age: " & HH_MEMB_ARRAY(i).age
				    DropListBox 180, y_pos, 110, 45, "03 - GED Or Equiv"+chr(9)+"13 - English As A 2nd Language", HH_MEMB_ARRAY(i).school_type
				    DropListBox 295, y_pos, 60, 45, schl_status_droplist, HH_MEMB_ARRAY(i).school_status
				    ComboBox 360, y_pos, 115, 45, schl_ver_droplist & chr(9)+HH_MEMB_ARRAY(i).school_verif, HH_MEMB_ARRAY(i).school_verif
					y_pos = y_pos + 20
				End If
			next
			If y_pos = start_pos Then
				Text 20, y_pos + 5, 200, 10, "No GED or ELL students known on this case."
				y_pos = y_pos + 20
			End if
			PushButton 355, ged_ell_btn_pos, 120, 13, "Add GED/ELL Student", add_ged_ell_student



		End If
		If page_display = show_q_4 Then										'THIS IS THE INFORMATION FOR PAGE  ----------------------------------------------------------------------------------------------------------------------------------
			Text 508, 72, 60, 13, "Q. 4"

			Text 5, 10, 305, 10, "^^1 - Enter the answers listed on the actual CAF form for Q4 into the 'Answer on the CAF' field."
			Text 15, 25, 335, 20, "Q. 4. Is anyone in your household temporarily not living in your home? (example: vacation, foster care, treatment, hospital job search)"
			Text 370, 25, 65, 10, "Answer on the CAF"
			DropListBox 435, 20, 40, 45, caf_answer_droplist, q4_caf_answer
			Text 5, 50, 35, 10, "^^2 - ASK - "
			Text 40, 50, 405, 20, "'Is there anyone who typically lives at home and is currently living elsewhere? Common examples are someone away for vacation, job search, but could also include treatment, hospital stay, or even foster care.'"
			Text 40, 80, 70, 10, "Confirm CAF Answer"
			ComboBox 110, 75, 365, 45, "", q4_confirm_caf_answer
			Text 40, 100, 245, 10, "Based on Information Provided, Are there individuals Temporarily absent?"
			DropListBox 285, 95, 60, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes", temp_absent_yn
		    GroupBox 10, 120, 465, 65, "^^3 - If YES to anyone Temporary Absent - ASK - the person information, where they are, and the dates they left and were expected to return."
		    Text 20, 135, 450, 10, "Person Absent --------------------------------------------------Where Living -------------------------------------------------------Left Date ----------------Expected Return Date"
			y_pos = 150
			For gone_membs = 0 to UBound(ALL_TEMP_ABSENCE, 2)
				ComboBox 20, y_pos, 140, 45, memb_droplist, ALL_TEMP_ABSENCE(temp_abs_person, gone_membs)
			    ComboBox 170, y_pos, 145, 45, "", ALL_TEMP_ABSENCE(temp_abs_where, gone_membs)
			    EditBox 325, y_pos, 55, 15, ALL_TEMP_ABSENCE(temp_abs_left_date, gone_membs)
			    EditBox 390, y_pos, 55, 15, ALL_TEMP_ABSENCE(temp_abs_ret_date, gone_membs)
				y_pos = y_pos + 20
			Next

			PushButton 345, y_pos, 100, 13, "Add Another Absent Person", add_another_absent_pers_btn
			y_pos = y_pos + 25
		    Text 5, y_pos, 210, 10, "^^4 - If YES to anyone Temporary Absent - EXPLAIN TO CLIENT:"
			y_pos = y_pos + 15
		    Text 20, y_pos, 455, 65, "ENTER TEMP ABSENCE POLICY HERE"
		End If
		If page_display = show_q_5 Then										'THIS IS THE INFORMATION FOR PAGE  ----------------------------------------------------------------------------------------------------------------------------------
			Text 508, 87, 60, 13, "Q. 5"

			Text 5, 10, 305, 10, "^^1 - Enter the answers listed on the actual CAF form for Q5 into the 'Answer on the CAF' field."
		    Text 20, 25, 335, 20, "Q. 5. Is anyone blind, or does anyone have a physical or mental health condition that limit the ability to work or perform daily activities?"
		    Text 370, 30, 65, 10, "Answer on the CAF"
		    DropListBox 435, 25, 40, 45, caf_answer_droplist, q5_caf_answer
		    Text 5, 50, 35, 10, "^^2 - ASK - "
		    Text 40, 50, 280, 10, "'Is anyone in the household disabled or have a physical or mental health condition?"
		    Text 40, 70, 70, 10, "Confirm CAF Answer"
		    ComboBox 110, 65, 365, 45, "", q5_confirm_caf_answer
		    Text 40, 90, 230, 10, "Based on Information Provided, Is anyone in the household Disabled?"
		    DropListBox 275, 85, 60, 45, "Select One..."+chr(9)+"No"+chr(9)+"Yes", temp_absent_yn
		    Text 5, 110, 300, 10, "^^3 - REVIEW information from MAXIS with client about known disabilities:"

			y_pos = 125

			For i = 0 to UBOUND(HH_MEMB_ARRAY, 1)
				If HH_MEMB_ARRAY(i).disa_exists = TRUE Then
					' Text 20, y_pos, 130, 10, "MEMB " & HH_MEMB_ARRAY(i).ref_number & " - " & HH_MEMB_ARRAY(i).full_name
					GroupBox 20, y_pos, 460, 60, "MEMB " & HH_MEMB_ARRAY(i).ref_number & " - " & HH_MEMB_ARRAY(i).full_name
				    ' Text 200, y_pos, 95, 10, "DISA End Date: " & HH_MEMB_ARRAY(i).disa_end_date
					Text 200, y_pos, 110, 10, "DISA: " & HH_MEMB_ARRAY(i).disa_detail
				    Text 315, y_pos, 45, 10, "DISA Review:"
				    DropListBox 365, y_pos - 5, 110, 45, "Select One..."+chr(9)+"DISA Ended"+chr(9)+"DISA Needs Verif"+chr(9)+"DISA Continues", HH_MEMB_ARRAY(i).disa_review
					y_pos = y_pos + 15
				    ' Text 35, y_pos, 110, 10, "DISA: " & HH_MEMB_ARRAY(i).disa_detail
					Text 25, y_pos, 70, 10, "MOF: MOF On File:"
					DropListBox 90, y_pos - 5, 85, 45, "Select One..."+chr(9)+"On File"+chr(9)+"Needed"+chr(9)+"Requested"+chr(9)+"Attached"+chr(9)+"Not Needed", HH_MEMB_ARRAY(i).mof_file
					Text 190, y_pos, 115, 10, "if received, Certification End Date:"
					EditBox 305, y_pos - 5, 50, 15, HH_MEMB_ARRAY(i).mof_end_date
					y_pos = y_pos + 15
					Text 25, y_pos, 95, 10, "DISA End Date: " & HH_MEMB_ARRAY(i).disa_end_date
					Text 145, y_pos, 95, 10, "DISA Cert End Date: " & HH_MEMB_ARRAY(i).disa_cert_end_date
					PushButton 335, y_pos - 3, 140, 13, "Update DISA Information for this MEMBER", HH_MEMB_ARRAY(i).button_two
					y_pos = y_pos + 15
				    Text 25, y_pos, 65, 10, "IAAs: IAAs On File:"
				    DropListBox 90, y_pos - 5, 85, 45, "Select One..."+chr(9)+"On File"+chr(9)+"Needed"+chr(9)+"Requested"+chr(9)+"Attached"+chr(9)+"Not Needed", HH_MEMB_ARRAY(i).iaa_file
				    Text 190, y_pos, 95, 10, "if received, Received Date:"
				    EditBox 280, y_pos - 5, 50, 15, HH_MEMB_ARRAY(i).iaa_received_date
				    CheckBox 335, y_pos, 140, 10, "Check here if IAAs are signed Correctly", HH_MEMB_ARRAY(i).iaa_complete
					y_pos = y_pos + 20
				End If
			Next
			If y_pos = 125 Then Text 20, 145, 400, 10, "There is know DISA information in MAXIS or added."

		    PushButton 345, y_pos, 130, 13, "Add New DISA for a Known Member", new_disa_btn

		End If
		If page_display = show_q_6 Then										'THIS IS THE INFORMATION FOR PAGE  ----------------------------------------------------------------------------------------------------------------------------------
			Text 508, 102, 60, 13, "Q. 6"

			Text 5, 10, 305, 10, "^^1 - Enter the answers listed on the actual CAF form for Q6 into the 'Answer on the CAF' field."
		    Text 20, 25, 335, 10, "Q. 6. Is anyone unable to work for reasons other than illness or disability?"
		    Text 370, 25, 65, 10, "Answer on the CAF"
		    DropListBox 435, 20, 40, 45, caf_answer_droplist, q6_caf_answer
		    Text 5, 45, 35, 10, "^^2 - ASK - "
		    Text 40, 45, 280, 10, "'Is anyone in the household disabled or have a physical or mental health condition?"
		    Text 40, 65, 70, 10, "Confirm CAF Answer"
		    ComboBox 110, 60, 365, 45, "", q6_confirm_caf_answer
		    Text 5, 85, 385, 10, "^^3 - If YES (based on above detail if the client indicates someone is unable to work) - ASK client to EXPLAIN in detail"

			y_pos = 100
			For each_note = 0 to UBound(NON_DISA_UNABLE_TO_WORK, 2)
				Text 20, y_pos, 85, 10, "Member Unable to Work"
				y_pos = y_pos + 10
				DropListBox 20, y_pos, 150, 45, memb_droplist, NON_DISA_UNABLE_TO_WORK(pers_unable_to_work, each_note)
				Text 185, y_pos + 5, 35, 10, "Start Date:"
				EditBox 220, y_pos, 50, 15, NON_DISA_UNABLE_TO_WORK(unable_to_work_start_date, each_note)
				Text 285, y_pos + 5, 40, 10, "Verification:"
				DropListBox 330, y_pos, 145, 40, "", NON_DISA_UNABLE_TO_WORK(unable_to_work_verif, each_note)
				y_pos = y_pos + 20
				Text 35, y_pos + 5, 30, 10, "Reason:"
				ComboBox 65, y_pos, 410, 45, "Select or Type"+chr(9)+"Care of a Child < 6"+chr(9)+"Care of a Child 6 or Over"+chr(9)+"Care of an Elderly Person"+chr(9)+"Care of a Disabled Person"+chr(9)+"Lack Access to facilities required for employment", NON_DISA_UNABLE_TO_WORK(unable_to_work_reason, each_note)
				y_pos = y_pos + 20
				Text 35, y_pos + 5, 100, 10, "Is this an ABAWD Exemption?"
				DropListBox 135, y_pos, 40, 45, "No"+chr(9)+"Yes", NON_DISA_UNABLE_TO_WORK(unable_to_work_abawd_yn, each_note)
				DropListBox 180, y_pos, 295, 45, "If YES - what is ABAWD Exemption this may meet?", NON_DISA_UNABLE_TO_WORK(unable_to_work_abawd_type, each_note)
				y_pos = y_pos + 20
				Text 35, y_pos + 5, 135, 10, "Does this meed MFIP Extention Criteria?"
				DropListBox 175, y_pos, 40, 45, "No"+chr(9)+"Yes", NON_DISA_UNABLE_TO_WORK(unable_to_work_mfip_yn, each_note)
				DropListBox 220, y_pos, 255, 45, "If YES - what exemption does this meet?", NON_DISA_UNABLE_TO_WORK(unable_to_work_mfip_type, each_note)
				y_pos = y_pos + 20
			Next
			PushButton 365, y_pos, 110, 10, "Add MEMBER Unable to Work", add_unable_tp_work_memb_btn

		End If
		If page_display = show_q_7 Then										'THIS IS THE INFORMATION FOR PAGE  ----------------------------------------------------------------------------------------------------------------------------------
			Text 508, 117, 60, 13, "Q. 7"

			Text 5, 10, 305, 10, "^^1 - Enter the answers listed on the actual CAF form for Q5 into the 'Answer on the CAF' field."
			Text 20, 25, 335, 20, "Q. 7. In the last 60 days did anyone in the household: Stop working or quit? Refuse a job offer? Ask to work fewwer hours? Go on strike?"
			Text 370, 30, 65, 10, "Answer on the CAF"
			DropListBox 435, 25, 40, 45, caf_answer_droplist, q7_caf_answer
			Text 5, 50, 35, 10, "^^2 - ASK - "
			Text 40, 50, 280, 10, "'Is anyone in the household disabled or have a physical or mental health condition?"
			Text 40, 70, 70, 10, "Confirm CAF Answer"
			ComboBox 110, 65, 365, 45, "", q7_confirm_caf_answer
			Text 5, 90, 315, 10, "^^3 - CONFIRM - Review details of known STWK or add details if not currently known."

			grp_len = 180

			If case_has_stwk_panel = FALSE Then
				GroupBox 50, 105, 425, 24, "STWK on CASE " & MAXIS_case_number
				Text 75, 115, 250, 10, "NO STWK PANELS OR OTHER STWK DETAILS ON THIS CASE"
			Else
				GroupBox 50, 105, 425, grp_len, "STWK for " & HH_MEMB_ARRAY(stwk_selected).full_name
				Text 75, 125, 165, 10, "Employer: " & HH_MEMB_ARRAY(stwk_selected).stwk_employer
				Text 75, 140, 165, 10, "Work Stop Date: " & HH_MEMB_ARRAY(stwk_selected).stwk_work_stop_date
				Text 275, 140, 165, 10, "Verif: " & HH_MEMB_ARRAY(stwk_selected).stwk_verification
				Text 75, 155, 165, 10, "Income Stop Date: " & HH_MEMB_ARRAY(stwk_selected).stwk_income_stop_date
				Text 275, 155, 165, 10, "Refused Empl: " & HH_MEMB_ARRAY(stwk_selected).stwk_refused_employment
				Text 75, 180, 165, 10, "Voluntary Quit: " & HH_MEMB_ARRAY(stwk_selected).stwk_vol_quit
				Text 275, 180, 165, 10, "Refused Empl Date: " & HH_MEMB_ARRAY(stwk_selected).stwk_refused_employment_date
				Text 75, 205, 165, 10, "Good Cause: Cash: " & HH_MEMB_ARRAY(stwk_selected).stwk_cash_good_cause_yn & "   GRH: " & HH_MEMB_ARRAY(stwk_selected).stwk_grh_good_cause_yn & "   GRH: " & HH_MEMB_ARRAY(stwk_selected).stwk_snap_good_cause_yn
				Text 75, 230, 165, 10, "FS PWE: " & HH_MEMB_ARRAY(stwk_selected).stwk_snap_pwe
				Text 75, 255, 165, 10, "MA-EPD Extension: " & HH_MEMB_ARRAY(stwk_selected).stwk_ma_epd_extension

				PushButton 55, 270, 125, 10, "Update STWK Information", stwk_info_change
				PushButton 410, 270, 60, 10, "NEXT MEMB", next_stwk_btn

				btn_pos = 110
				For i = 0 to Ubound(HH_MEMB_ARRAY, 1)
					If i = stwk_selected Then
						Text 9, btn_pos+1, 40, 10, "MEMB " & HH_MEMB_ARRAY(i).ref_number
					ElseIf HH_MEMB_ARRAY(i).stwk_exists = TRUE Then
						PushButton 5, btn_pos, 40, 10, "MEMB " & HH_MEMB_ARRAY(i).ref_number, HH_MEMB_ARRAY(i).button_one
					End If
					btn_pos = btn_pos + 10
				Next
				btn_pos = btn_pos + 20
			End If

			PushButton 5, 290, 150, 10, "ADD STWK FOR ANOTHER MEMBER", add_new_stwk_btn


		End If
		If page_display = show_q_8 Then										'THIS IS THE INFORMATION FOR PAGE  ----------------------------------------------------------------------------------------------------------------------------------
			Text 508, 132, 60, 13, "Q. 8"

			Text 5, 10, 305, 10, "^^1 - Enter the answers listed on the actual CAF form for Q5 into the 'Answer on the CAF' field."
			Text 20, 25, 335, 20, "Q. 8. Has anyone in the household had a job or been self-employed in the past 12 months?"
			Text 370, 30, 65, 10, "Answer on the CAF"
			DropListBox 435, 25, 40, 45, caf_answer_droplist, q8_caf_answer
			Text 5, 50, 35, 10, "^^2 - ASK - "
			Text 40, 50, 280, 10, "'Is anyone in the household disabled or have a physical or mental health condition?"
			Text 40, 70, 70, 10, "Confirm CAF Answer"
			ComboBox 110, 65, 365, 45, "", q8_confirm_caf_answer


			Text 5, 90, 305, 10, "^^3 - Enter the answers listed on the actual CAF form for Q5 into the 'Answer on the CAF' field."
			Text 20, 105, 335, 20, "Q. 8a. FOR SNAP ONLY: Has anyone in the household had a job or been self-employed in the past 36 months?"
			Text 370, 110, 65, 10, "Answer on the CAF"
			DropListBox 435, 108, 40, 45, caf_answer_droplist, q8a_caf_answer
			Text 5, 130, 35, 10, "^^4 - ASK - "
			Text 40, 130, 280, 10, "'Is anyone in the household disabled or have a physical or mental health condition?"
			Text 40, 150, 70, 10, "Confirm CAF Answer"
			ComboBox 110, 145, 365, 45, "", q8a_confirm_caf_answer


		End If
		If page_display = show_q_9 Then										'THIS IS THE INFORMATION FOR PAGE  ----------------------------------------------------------------------------------------------------------------------------------
			Text 508, 147, 60, 13, "Q. 9"

			'NOTES ABOUT ENTERING JOBS
			'INTV script will not update JOBS because we cannot fit EI budgeting into this scrpit.(except maybe create a new panel)
			'EI Budgeting can read from ITVW case note to gather initial JOBS details that were captured in the note
			'THE PURPOSE HERE IS TO ASK THE RIGHT QUESTIONS

			Text 5, 10, 305, 10, "^^1 - Enter the answers listed on the actual CAF form for Q5 into the 'Answer on the CAF' field."
			Text 20, 25, 335, 20, "Q. 9. Does anyone in the household have a job or expect to get income from a job this month or next month? (Include income from Work Study and paid scholarships. Include free benefits or reduced expenses received for work (shelter, food, clothing, etc.)"
			Text 370, 30, 65, 10, "Answer on the CAF"
			DropListBox 435, 25, 40, 45, caf_answer_droplist, q9_caf_answer
			Text 5, 50, 35, 10, "^^2 - ASK - "
			Text 40, 50, 280, 10, "'Is anyone in the household disabled or have a physical or mental health condition?"
			Text 40, 70, 70, 10, "Confirm CAF Answer"
			ComboBox 110, 65, 365, 45, "", q9_confirm_caf_answer

		End If
		If page_display = show_q_10 Then										'THIS IS THE INFORMATION FOR PAGE  ----------------------------------------------------------------------------------------------------------------------------------
			Text 507, 162, 60, 13, "Q. 10"

			Text 5, 10, 305, 10, "^^1 - Enter the answers listed on the actual CAF form for Q5 into the 'Answer on the CAF' field."
			Text 20, 25, 335, 20, "Q. 10. Is anyone in the household self-employed or does anyone expect to get income from self-employment this month or next month?"
			Text 370, 30, 65, 10, "Answer on the CAF"
			DropListBox 435, 25, 40, 45, caf_answer_droplist, q10_caf_answer
			Text 5, 50, 35, 10, "^^2 - ASK - "
			Text 40, 50, 280, 10, "'Is anyone in the household disabled or have a physical or mental health condition?"
			Text 40, 70, 70, 10, "Confirm CAF Answer"
			ComboBox 110, 65, 365, 45, "", q10_confirm_caf_answer

		End If
		If page_display = show_q_11 Then										'THIS IS THE INFORMATION FOR PAGE  ----------------------------------------------------------------------------------------------------------------------------------
			Text 507, 177, 60, 13, "Q. 11"

			Text 5, 10, 305, 10, "^^1 - Enter the answers listed on the actual CAF form for Q5 into the 'Answer on the CAF' field."
			Text 20, 25, 335, 20, "Q. 11. Do you expect any changes in income, expenses or work hours?"
			Text 370, 30, 65, 10, "Answer on the CAF"
			DropListBox 435, 25, 40, 45, caf_answer_droplist, q11_caf_answer
			Text 5, 50, 35, 10, "^^2 - ASK - "
			Text 40, 50, 280, 10, "'Is anyone in the household disabled or have a physical or mental health condition?"
			Text 40, 70, 70, 10, "Confirm CAF Answer"
			ComboBox 110, 65, 365, 45, "", q11_confirm_caf_answer


		End If
		If page_display = show_q_12 Then										'THIS IS THE INFORMATION FOR PAGE  ----------------------------------------------------------------------------------------------------------------------------------
			' MsgBox second_page_display
			Text 507, 192, 60, 13, "Q. 12"

			Text 5, 10, 195, 10, "^^1 - Enter the answers listed on the actual CAF from Q12"
			Text 25, 30, 85, 10, "Social Security (RSDI)"
		    DropListBox 110, 25, 40, 45, caf_answer_droplist, rsdi_caf_answer
		    Text 155, 30, 5, 10, "$"
		    EditBox 165, 25, 40, 15, rsdi_caf_amt
		    Text 230, 30, 120, 10, "Supplemental Security Income (SSI)"
		    DropListBox 355, 25, 40, 45, caf_answer_droplist, ssi_caf_answer
		    Text 400, 30, 5, 10, "$"
		    EditBox 410, 25, 40, 15, ssi_caf_amt
		    Text 25, 45, 85, 10, "Veteran Beneftis (VA)"
		    DropListBox 110, 40, 40, 45, caf_answer_droplist, va_caf_answer
		    Text 155, 45, 5, 10, "$"
		    EditBox 165, 40, 40, 15, va_caf_amt
		    Text 230, 45, 120, 10, "Unemployment Insurance"
		    DropListBox 355, 40, 40, 45, caf_answer_droplist, ui_caf_answer
		    Text 400, 45, 5, 10, "$"
		    EditBox 410, 40, 40, 15, ui_caf_amt
		    Text 25, 60, 85, 10, "Workers' Compensation"
		    DropListBox 110, 55, 40, 45, caf_answer_droplist, wc_caf_answer
		    Text 155, 60, 5, 10, "$"
		    EditBox 165, 55, 40, 15, wc_caf_amt
		    Text 230, 60, 120, 10, "Retirement Benefits"
		    DropListBox 355, 55, 40, 45, caf_answer_droplist, ret_caf_answer
		    Text 400, 60, 5, 10, "$"
		    EditBox 410, 55, 40, 15, ret_caf_amt
		    Text 25, 75, 85, 10, "Tribal Payments"
		    DropListBox 110, 70, 40, 45, caf_answer_droplist, tribal_caf_answer
		    Text 155, 75, 5, 10, "$"
		    EditBox 165, 70, 40, 15, tribal_caf_amt
		    Text 230, 75, 120, 10, "Child Support or Spousal Support"
		    DropListBox 355, 70, 40, 45, caf_answer_droplist, cs_caf_answer
		    Text 400, 75, 5, 10, "$"
		    EditBox 410, 70, 40, 15, cs_caf_amt
		    Text 25, 90, 85, 10, "Other Unearned Income"
		    DropListBox 110, 85, 40, 45, caf_answer_droplist, other_unea_caf_answer
		    Text 155, 90, 5, 10, "$"
		    EditBox 165, 85, 40, 15, other_unea_caf_amt

			rsdi_count = 0
			ssi_count = 0
			va_count = 0
			ui_count = 0
			wc_count = 0
			retirement_count = 0
			tribal_count = 0
			cs_count = 0
			ss_count = 0
			other_UNEA_count = 0
			If unea_found = TRUE Then
				for i = start_count to stop_count
					' MsgBox "HERE"
					INCOME_ARRAY(i).read_member_name
					INCOME_ARRAY(i).read_unea_panel

					If INCOME_ARRAY(i).income_type_code = "01" OR INCOME_ARRAY(i).income_type_code = "02" Then rsdi_count = rsdi_count + 1
					If INCOME_ARRAY(i).income_type_code = "03" Then ssi_count = ssi_count + 1
					If INCOME_ARRAY(i).income_type_code = "15" Then wc_count = wc_count + 1
					If INCOME_ARRAY(i).income_type_code = "14" Then ui_count = ui_count + 1
					If INCOME_ARRAY(i).income_type_code = "11" OR INCOME_ARRAY(i).income_type_code = "12" OR INCOME_ARRAY(i).income_type_code = "13" OR INCOME_ARRAY(i).income_type_code = "38" Then va_count = va_count + 1
					If INCOME_ARRAY(i).income_type_code = "16" OR INCOME_ARRAY(i).income_type_code = "17" Then retirement_count = retirement_count + 1
					If INCOME_ARRAY(i).income_type_code = "46" OR INCOME_ARRAY(i).income_type_code = "47" Then tribal_count = tribal_count + 1
					If INCOME_ARRAY(i).income_type_code = "08" OR INCOME_ARRAY(i).income_type_code = "36" OR INCOME_ARRAY(i).income_type_code = "39" OR INCOME_ARRAY(i).income_type_code = "43" OR INCOME_ARRAY(i).income_type_code = "45" Then cs_count = cs_count + 1
					If INCOME_ARRAY(i).income_type_code = "35" OR INCOME_ARRAY(i).income_type_code = "37" OR INCOME_ARRAY(i).income_type_code = "40" Then ss_count = ss_count + 1
					If INCOME_ARRAY(i).income_type_code = "06" OR INCOME_ARRAY(i).income_type_code = "18" OR INCOME_ARRAY(i).income_type_code = "19" OR INCOME_ARRAY(i).income_type_code = "20" OR INCOME_ARRAY(i).income_type_code = "21" OR INCOME_ARRAY(i).income_type_code = "22" OR INCOME_ARRAY(i).income_type_code = "23" OR INCOME_ARRAY(i).income_type_code = "24" OR INCOME_ARRAY(i).income_type_code = "25" OR INCOME_ARRAY(i).income_type_code = "26" OR INCOME_ARRAY(i).income_type_code = "27" OR INCOME_ARRAY(i).income_type_code = "28" OR INCOME_ARRAY(i).income_type_code = "29" OR INCOME_ARRAY(i).income_type_code = "30" OR INCOME_ARRAY(i).income_type_code = "31" OR INCOME_ARRAY(i).income_type_code = "44" OR INCOME_ARRAY(i).income_type_code = "48" OR INCOME_ARRAY(i).income_type_code = "49" Then other_UNEA_count = other_UNEA_count + 1
				next
			End If

			' Text 25, 100, 400, 20, "Use the Buttons below to ask about details for each type of unearned income. The numbers on the buttons indicate how many panels of each type of income is known."

			If second_page_display = main_unea Then
				Text 5, 115, 35, 10, "^^2 - ASK - "
			    Text 40, 115, 280, 10, "'Has anyone in the household applied for or receive any ..."
				Text 100, 135, 75, 10, "RSDI - Social Security"
			    DropListBox 200, 130, 180, 45, "", rsdi_confirm_response
			    Text 100, 155, 70, 10, "SSI - Social Security"
			    DropListBox 200, 150, 180, 45, "", ssi_confirm_response
			    Text 100, 175, 70, 10, "Veteran Benefits (VA)"
			    DropListBox 200, 170, 180, 45, "", va_confirm_response
			    Text 100, 195, 65, 10, "Unemployment (UI)"
			    DropListBox 200, 190, 180, 45, "", ui_confirm_response
			    Text 100, 215, 100, 10, "Workers' Compensation (WC)"
			    DropListBox 200, 210, 180, 45, "", wc_confirm_response
			    Text 100, 235, 70, 10, "Retirement Benefits"
			    DropListBox 200, 230, 180, 45, "", ret_confirm_response
			    Text 100, 255, 55, 10, "Tribal Payments"
			    DropListBox 200, 250, 180, 45, "", tribal_confirm_response
			    Text 100, 275, 45, 10, "Child Support"
			    DropListBox 200, 270, 180, 45, "", cs_confirm_response
			    Text 100, 295, 55, 10, "Spousal Support"
			    DropListBox 200, 290, 180, 45, "", ss_confirm_response
			    Text 100, 315, 45, 10, "Other UNEA"
			    DropListBox 200, 310, 180, 45, "", other_unea_confirm_response

				Text 48, 332, 70, 15, "Main"
			End If

			If second_page_display = rsdi_unea Then								'=====================================================================================   UNEA - RSDI
				GroupBox 100, 125, 380, 220, "RSDI Income"

				x_pos = 110
				first_rsdi = TRUE
				for i = 0 to UBound(INCOME_ARRAY, 1)
					If INCOME_ARRAY(i).panel_name = "UNEA" Then
						If INCOME_ARRAY(i).income_type_code = "01" OR INCOME_ARRAY(i).income_type_code = "02" Then

							show_rsdi = FALSE
							for j = 0 to UBound(HH_MEMB_ARRAY, 1)
								If HH_MEMB_ARRAY(j).ref_number = INCOME_ARRAY(i).member_ref Then
									If first_rsdi = TRUE and memb_to_match = "" Then
										Text x_pos + 5, 331, 35, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number
										first_rsdi = FALSE
										show_rsdi = TRUE
									ElseIf memb_to_match = HH_MEMB_ARRAY(j).ref_number Then
										Text x_pos + 5, 331, 35, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number
										show_rsdi = TRUE
									Else
										PushButton x_pos, 330, 40, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number, HH_MEMB_ARRAY(j).button_one
									End If
									x_pos = x_pos + 40
								End If
							Next

							If show_rsdi = TRUE Then
								Text 110, 135, 160, 10, "HH Member: " & INCOME_ARRAY(i).member
							    Text 280, 135, 105, 10, "Claim Number: " & INCOME_ARRAY(i).claim_number
								Text 110, 150, 270, 10, "RSDI Income Type: " & INCOME_ARRAY(i).income_type
								Text 110, 165, 120, 10, "Date Most Recent Income Received:"
							    EditBox 235, 160, 45, 15, INCOME_ARRAY(i).most_recent_pay_date
							    Text 290, 165, 130, 10, "How Much was the Most Recent Check:"
							    EditBox 425, 160, 50, 15, INCOME_ARRAY(i).most_recent_pay_amt
							    Text 110, 180, 110, 10, "Known Monthly Income: $" & INCOME_ARRAY(i).prosp_pay_total
							    Text 110, 195, 355, 20, "If the known income and the most recent income received does not match, press the 'Update RSDI Information' button to clarify the income to budget and other details."
							    Text 110, 215, 90, 10, "Start Date: " & INCOME_ARRAY(i).income_start_date
							    Text 215, 215, 85, 10, "End Date: " & INCOME_ARRAY(i).income_end_date
							    Text 110, 230, 50, 10, "Verification"
							    Text 240, 230, 50, 10, "Verif Info"
							    DropListBox 110, 240, 125, 45, unea_verif_droplist, INCOME_ARRAY(i).income_verification
							    EditBox 240, 240, 235, 15, INCOME_ARRAY(i).verif_explaination
							    Text 110, 260, 50, 10, "Income Notes"
							    EditBox 110, 270, 365, 15, INCOME_ARRAY(i).income_notes
							    Text 110, 295, 70, 10, "Review of Income"
							    ComboBox 110, 305, 365, 45, "", INCOME_ARRAY(i).income_review

							    PushButton 390, 135, 85, 15, "Update RSDI Information", update_rsdi_info_btn
							End If
						End If
					End If
				next
				If rsdi_count = 0 Then
					Text 110, 140, 355, 20, "There are no RSDI panels known in MAXIS and no additional RSDI Income information has been added."
				End If
				PushButton 385, 347, 95, 13, "Add RSDI Information", add_another_rsdi_unea_btn

				Text 42, 132, 45, 15, "RSDI - " & rsdi_count
			End If
			If second_page_display = ssi_unea Then								'=====================================================================================   UNEA - SSI
				GroupBox 100, 125, 380, 220, "SSI Income"

				x_pos = 110
				first_ssi = TRUE
				for i = 0 to UBound(INCOME_ARRAY, 1)
					If INCOME_ARRAY(i).panel_name = "UNEA" Then
						If INCOME_ARRAY(i).income_type_code = "03" Then
							show_ssi = FALSE
							for j = 0 to UBound(HH_MEMB_ARRAY, 1)
								If HH_MEMB_ARRAY(j).ref_number = INCOME_ARRAY(i).member_ref Then
									If first_ssi = TRUE and memb_to_match = "" Then
										Text x_pos + 5, 331, 35, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number
										first_ssi = FALSE
										show_ssi = TRUE
									ElseIf memb_to_match = HH_MEMB_ARRAY(j).ref_number Then
										Text x_pos + 5, 331, 35, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number
										show_ssi = TRUE
									Else
										PushButton x_pos, 330, 40, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number, HH_MEMB_ARRAY(j).button_one
									End If
									x_pos = x_pos + 40
								End If
							Next

							If show_ssi = TRUE Then
								Text 110, 140, 160, 10, "HH Member: " & INCOME_ARRAY(i).member
							    Text 280, 140, 105, 10, "Claim Number: " & INCOME_ARRAY(i).claim_number
								Text 110, 160, 120, 10, "Date Most Recent Income Received:"
							    EditBox 235, 155, 45, 15, INCOME_ARRAY(i).most_recent_pay_date
							    Text 290, 160, 130, 10, "How Much was the Most Recent Check:"
							    EditBox 425, 155, 50, 15, INCOME_ARRAY(i).most_recent_pay_amt
							    Text 110, 175, 110, 10, "Known Monthly Income: $" & INCOME_ARRAY(i).prosp_pay_total
							    Text 110, 190, 355, 20, "If the known income and the most recent income received does not match, press the 'Update SSI Information' button to clarify the income to budget and other details."
							    Text 110, 210, 90, 10, "Start Date: " & INCOME_ARRAY(i).income_start_date
							    Text 215, 210, 85, 10, "End Date: " & INCOME_ARRAY(i).income_end_date
							    Text 110, 230, 50, 10, "Verification"
							    Text 240, 230, 50, 10, "Verif Info"
							    DropListBox 110, 240, 125, 45, unea_verif_droplist, INCOME_ARRAY(i).income_verification
							    EditBox 240, 240, 235, 15, INCOME_ARRAY(i).verif_explaination
							    Text 110, 260, 50, 10, "Income Notes"
							    EditBox 110, 270, 365, 15, INCOME_ARRAY(i).income_notes
							    Text 110, 295, 70, 10, "Review of Income"
							    ComboBox 110, 305, 365, 45, "", INCOME_ARRAY(i).income_review

							    PushButton 390, 135, 85, 15, "Update SSI Information", update_ssi_info_btn
							End If
						End If
					End If
				next
				If ssi_count = 0 Then
					Text 110, 140, 355, 20, "There are no SSI panels known in MAXIS and no additional SSI Income information has been added."
				End If
				PushButton 385, 347, 95, 13, "Add SSI Information", add_another_ssi_unea_btn

				Text 45, 152, 42, 15, "SSI - " & ssi_count
			End If
			If second_page_display = va_unea Then								'=====================================================================================   UNEA - VETERANS INCOME
				GroupBox 100, 125, 380, 220, "VA Income"

				x_pos = 110
				first_va = TRUE
				for i = 0 to UBound(INCOME_ARRAY, 1)
					If INCOME_ARRAY(i).panel_name = "UNEA" Then
						If INCOME_ARRAY(i).income_type_code = "11" OR INCOME_ARRAY(i).income_type_code = "12" OR INCOME_ARRAY(i).income_type_code = "13" OR INCOME_ARRAY(i).income_type_code = "38" Then

							show_va = FALSE
							for j = 0 to UBound(HH_MEMB_ARRAY, 1)
								If HH_MEMB_ARRAY(j).ref_number = INCOME_ARRAY(i).member_ref Then
									If first_va = TRUE and memb_to_match = "" Then
										Text x_pos + 5, 331, 35, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number
										first_va = FALSE
										show_va = TRUE
									ElseIf memb_to_match = HH_MEMB_ARRAY(j).ref_number Then
										Text x_pos + 5, 331, 35, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number
										show_va = TRUE
									Else
										PushButton x_pos, 330, 40, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number, HH_MEMB_ARRAY(j).button_one
									End If
									x_pos = x_pos + 40
								End If
							Next
							If show_va = TRUE Then
								Text 110, 140, 160, 10, "HH Member: " & INCOME_ARRAY(i).member
							    Text 280, 140, 105, 10, "Claim Number: " & INCOME_ARRAY(i).claim_number
							    Text 110, 160, 120, 10, "Date Most Recent Income Received:"
							    EditBox 235, 155, 45, 15, INCOME_ARRAY(i).most_recent_pay_date
							    Text 290, 160, 130, 10, "How Much was the Most Recent Check:"
							    EditBox 425, 155, 50, 15, INCOME_ARRAY(i).most_recent_pay_amt
							    Text 110, 175, 110, 10, "Known Monthly Income: $" & INCOME_ARRAY(i).prosp_pay_total
							    Text 230, 175, 245, 10, "VA Income Type: " & INCOME_ARRAY(i).income_type
							    Text 110, 195, 75, 10, "Gross Monthly Income:"
							    EditBox 190, 190, 35, 15, INCOME_ARRAY(i).pay_gross
							    Text 235, 195, 70, 10, "Allowed Exclusions:"
							    EditBox 305, 190, 35, 15, INCOME_ARRAY(i).expenses_allowed
							    Text 355, 195, 85, 10, "Exclusions NOT Allowed:"
							    EditBox 440, 190, 35, 15, INCOME_ARRAY(i).expenses_not_allowed
							    Text 110, 210, 355, 10, "If the counted income is incorrect, press the Update Income button."
							    Text 110, 225, 90, 10, "Start Date: " & INCOME_ARRAY(i).income_start_date
							    Text 215, 225, 85, 10, "End Date: " & INCOME_ARRAY(i).income_end_date
							    Text 110, 240, 50, 10, "Verification"
							    Text 240, 240, 50, 10, "Verif Info"
							    DropListBox 110, 250, 125, 45, unea_verif_droplist, INCOME_ARRAY(i).income_verification
							    EditBox 240, 250, 235, 15, INCOME_ARRAY(i).verif_explaination
							    Text 110, 265, 50, 10, "Income Notes"
							    EditBox 110, 275, 365, 15, INCOME_ARRAY(i).income_notes
							    Text 110, 295, 70, 10, "Review of Income"
							    ComboBox 110, 305, 365, 45, "", INCOME_ARRAY(i).income_review

								PushButton 390, 135, 85, 15, "Update Income", update_va_info_btn
							End If
						End If
					End If
				next
				If va_count = 0 Then
					Text 110, 140, 355, 20, "There are no VA panels known in MAXIS and no additional VA Income information has been added."
				End If
				PushButton 385, 347, 95, 13, "Add VA Information", add_another_va_unea_btn

				Text 47, 172, 40, 15, "VA - " & va_count
			End If
			If second_page_display = ui_unea Then								'=====================================================================================   UNEA - UNEMPLOYMENT
				GroupBox 100, 125, 380, 220, "UI Income"

				x_pos = 110
				first_ui = TRUE
				for i = 0 to UBound(INCOME_ARRAY, 1)
					If INCOME_ARRAY(i).panel_name = "UNEA" Then
						If INCOME_ARRAY(i).income_type_code = "14" Then
							' MsgBox memb_droplist & vbNewLine & "~" & INCOME_ARRAY(i).member & "~" & vbNewLine & unea_verif_droplist & vbNewLine & "~" & INCOME_ARRAY(i).income_verification & "~" & vbNewLine & days_of_the_week_droplist & vbNewLine & "~" &  INCOME_ARRAY(i).pay_weekday & "~"

							show_ui = FALSE
							for j = 0 to UBound(HH_MEMB_ARRAY, 1)
								If HH_MEMB_ARRAY(j).ref_number = INCOME_ARRAY(i).member_ref Then
									If first_ui = TRUE and memb_to_match = "" Then
										Text x_pos + 5, 331, 35, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number
										first_ui = FALSE
										show_ui = TRUE
									ElseIf memb_to_match = HH_MEMB_ARRAY(j).ref_number Then
										Text x_pos + 5, 331, 35, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number
										show_ui = TRUE
									Else
										PushButton x_pos, 330, 40, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number, HH_MEMB_ARRAY(j).button_one
									End If
									x_pos = x_pos + 40
								End If
							Next
							If show_ui = TRUE Then
								Text 110, 140, 160, 10, "HH Member: " & INCOME_ARRAY(i).member
							    Text 280, 140, 105, 10, "Claim Number: " & INCOME_ARRAY(i).claim_number
							    Text 110, 160, 120, 10, "Date Most Recent Income Received:"
							    EditBox 235, 155, 45, 15, INCOME_ARRAY(i).most_recent_pay_date
							    Text 290, 160, 130, 10, "How Much was the Most Recent Check:"
							    EditBox 425, 155, 50, 15, INCOME_ARRAY(i).most_recent_pay_amt
							    Text 110, 175, 110, 10, "Known Weekly Income: $" & INCOME_ARRAY(i).prosp_average_pay
							    Text 110, 195, 75, 10, "Gross Weekly Income:"
							    EditBox 190, 190, 35, 15, INCOME_ARRAY(i).pay_gross
							    Text 235, 195, 70, 10, "Allowed Exclusions:"
							    EditBox 305, 190, 35, 15, INCOME_ARRAY(i).expenses_allowed
							    Text 355, 195, 85, 10, "Exclusions NOT Allowed:"
							    EditBox 440, 190, 35, 15, INCOME_ARRAY(i).expenses_not_allowed
							    Text 110, 210, 355, 10, "If the counted income is incorrect, press the Update Income button."
							    Text 110, 225, 90, 10, "Start Date: " & INCOME_ARRAY(i).income_start_date
							    Text 215, 225, 85, 10, "End Date: " & INCOME_ARRAY(i).income_end_date
								Text 360, 225, 50, 10, "Pay Weekday:"
							    DropListBox 415, 220, 60, 45, days_of_the_week_droplist, INCOME_ARRAY(i).pay_weekday
							    Text 110, 240, 50, 10, "Verification"
							    Text 240, 240, 50, 10, "Verif Info"
							    DropListBox 110, 250, 125, 45, unea_verif_droplist, INCOME_ARRAY(i).income_verification
							    EditBox 240, 250, 235, 15, INCOME_ARRAY(i).verif_explaination
							    Text 110, 265, 50, 10, "Income Notes"
							    EditBox 110, 275, 365, 15, INCOME_ARRAY(i).income_notes
							    Text 110, 295, 70, 10, "Review of Income"
							    ComboBox 110, 305, 365, 45, "", INCOME_ARRAY(i).income_review

							    PushButton 390, 135, 85, 15, "Update UI Information", update_ui_info_btn
							End If
						End If
					End If
				next
				If ui_count = 0 Then
					Text 110, 140, 355, 20, "There are no UI panels known in MAXIS and no additional UI Income information has been added."
				End If
				PushButton 385, 347, 95, 13, "Add UI Information", add_another_ui_unea_btn

				Text 48, 192, 18, 15, "UI - " & ui_count
			End If
			If second_page_display = wc_unea Then								'=====================================================================================   UNEA - WORKMANS COMP
				GroupBox 100, 125, 380, 220, "WC Income"

				x_pos = 110
				first_wc = TRUE
				for i = 0 to UBound(INCOME_ARRAY, 1)
					If INCOME_ARRAY(i).panel_name = "UNEA" Then
						If INCOME_ARRAY(i).income_type_code = "15" Then
							show_wc = FALSE
							for j = 0 to UBound(HH_MEMB_ARRAY, 1)
								If HH_MEMB_ARRAY(j).ref_number = INCOME_ARRAY(i).member_ref Then
									If first_wc = TRUE and memb_to_match = "" Then
										Text x_pos + 5, 331, 35, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number
										first_wc = FALSE
										show_wc = TRUE
									ElseIf memb_to_match = HH_MEMB_ARRAY(j).ref_number Then
										Text x_pos + 5, 331, 35, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number
										show_wc = TRUE
									Else
										PushButton x_pos, 330, 40, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number, HH_MEMB_ARRAY(j).button_one
									End If
									x_pos = x_pos + 40
								End If
							Next
							If show_wc = TRUE Then
								Text 110, 140, 160, 10, "HH Member: " & INCOME_ARRAY(i).member
								Text 280, 140, 105, 10, "Claim Number: " & INCOME_ARRAY(i).claim_number
								Text 110, 160, 120, 10, "Date Most Recent Income Received:"
								EditBox 235, 155, 45, 15, INCOME_ARRAY(i).most_recent_pay_date
								Text 290, 160, 130, 10, "How Much was the Most Recent Check:"
								EditBox 425, 155, 50, 15, INCOME_ARRAY(i).most_recent_pay_amt
								Text 110, 175, 110, 10, "Known Weekly Income: $" & INCOME_ARRAY(i).prosp_average_pay
								Text 110, 195, 75, 10, "Gross Weekly Income:"
								EditBox 190, 190, 35, 15, INCOME_ARRAY(i).pay_gross
								Text 235, 195, 70, 10, "Allowed Exclusions:"
								EditBox 305, 190, 35, 15, INCOME_ARRAY(i).expenses_allowed
								Text 355, 195, 85, 10, "Exclusions NOT Allowed:"
								EditBox 440, 190, 35, 15, INCOME_ARRAY(i).expenses_not_allowed
								Text 110, 210, 355, 10, "If the counted income is incorrect, press the Update Income button."
								Text 110, 225, 90, 10, "Start Date: " & INCOME_ARRAY(i).income_start_date
								Text 215, 225, 85, 10, "End Date: " & INCOME_ARRAY(i).income_end_date
								Text 360, 225, 50, 10, "Pay Weekday:"
								DropListBox 415, 220, 60, 45, days_of_the_week_droplist, INCOME_ARRAY(i).pay_weekday
								Text 110, 240, 50, 10, "Verification"
								Text 240, 240, 50, 10, "Verif Info"
								DropListBox 110, 250, 125, 45, unea_verif_droplist, INCOME_ARRAY(i).income_verification
								EditBox 240, 250, 235, 15, INCOME_ARRAY(i).verif_explaination
								Text 110, 265, 50, 10, "Income Notes"
								EditBox 110, 275, 365, 15, INCOME_ARRAY(i).income_notes
								Text 110, 295, 70, 10, "Review of Income"
								ComboBox 110, 305, 365, 45, "", INCOME_ARRAY(i).income_review

								PushButton 390, 135, 85, 15, "Update WC Information", update_wc_info_btn
							End If
						End If
					End If
				next
				If wc_count = 0 Then
					Text 110, 140, 355, 20, "There are no WC panels known in MAXIS and no additional WC Income information has been added."
				End If
				PushButton 385, 347, 95, 13, "Add WC Information", add_another_wc_unea_btn

				Text 46, 212, 41, 15, "WC - " & wc_count
			End If
			If second_page_display = ret_unea Then								'=====================================================================================   UNEA - RETIREMENT
				GroupBox 100, 125, 380, 220, "Retirement Income"

				x_pos = 110
				first_ri = TRUE
				for i = 0 to UBound(INCOME_ARRAY, 1)
					If INCOME_ARRAY(i).panel_name = "UNEA" Then
						If INCOME_ARRAY(i).income_type_code = "16" OR INCOME_ARRAY(i).income_type_code = "17" Then
							show_ri = FALSE
							for j = 0 to UBound(HH_MEMB_ARRAY, 1)
								If HH_MEMB_ARRAY(j).ref_number = INCOME_ARRAY(i).member_ref Then
									If first_ri = TRUE and memb_to_match = "" Then
										Text x_pos + 5, 331, 35, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number
										first_ri = FALSE
										show_ri = TRUE
									ElseIf memb_to_match = HH_MEMB_ARRAY(j).ref_number Then
										Text x_pos + 5, 331, 35, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number
										show_ri = TRUE
									Else
										PushButton x_pos, 330, 40, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number, HH_MEMB_ARRAY(j).button_one
									End If
									x_pos = x_pos + 40
								End If
							Next
							If show_ri = TRUE Then
								Text 110, 140, 160, 10, "HH Member: " & INCOME_ARRAY(i).member
								Text 280, 140, 105, 10, "Claim Number: " & INCOME_ARRAY(i).claim_number
								Text 110, 160, 120, 10, "Date Most Recent Income Received:"
								EditBox 235, 155, 45, 15, INCOME_ARRAY(i).most_recent_pay_date
								Text 290, 160, 130, 10, "How Much was the Most Recent Check:"
								EditBox 425, 155, 50, 15, INCOME_ARRAY(i).most_recent_pay_amt
								Text 110, 175, 110, 10, "Known Monthly Income: $" & INCOME_ARRAY(i).prosp_pay_total
								Text 230, 175, 245, 10, "Ret Income Type: " & INCOME_ARRAY(i).income_type
								Text 110, 195, 75, 10, "Gross Monthly Income:"
								EditBox 190, 190, 35, 15, INCOME_ARRAY(i).pay_gross
								Text 235, 195, 70, 10, "Allowed Exclusions:"
								EditBox 305, 190, 35, 15, INCOME_ARRAY(i).expenses_allowed
								Text 355, 195, 85, 10, "Exclusions NOT Allowed:"
								EditBox 440, 190, 35, 15, INCOME_ARRAY(i).expenses_not_allowed
								Text 110, 210, 355, 10, "If the counted income is incorrect, press the Update Income button."
								Text 110, 225, 90, 10, "Start Date: " & INCOME_ARRAY(i).income_start_date
								Text 215, 225, 85, 10, "End Date: " & INCOME_ARRAY(i).income_end_date
								Text 110, 240, 50, 10, "Verification"
								Text 240, 240, 50, 10, "Verif Info"
								DropListBox 110, 250, 125, 45, unea_verif_droplist, INCOME_ARRAY(i).income_verification
								EditBox 240, 250, 235, 15, INCOME_ARRAY(i).verif_explaination
								Text 110, 265, 50, 10, "Income Notes"
								EditBox 110, 275, 365, 15, INCOME_ARRAY(i).income_notes
								Text 110, 295, 70, 10, "Review of Income"
								ComboBox 110, 305, 365, 45, "", INCOME_ARRAY(i).income_review

								PushButton 390, 135, 85, 15, "Update Income", update_ri_info_btn
							End If
						End If
					End If
				next
				If retirement_count = 0 Then
					Text 110, 140, 355, 20, "There are no Retirement panels known in MAXIS and no additional Retirement Income information has been added."
				End If
				PushButton 385, 347, 95, 13, "Add Retirement Information", add_another_unea_btn

				Text 35, 232, 52, 15, "Retirement - " & retirement_count
			End If
			If second_page_display = tribal_unea Then								'=====================================================================================   UNEA - TRIBAL
				GroupBox 100, 125, 380, 220, "Tribal Income"

				x_pos = 110
				first_ti = TRUE
				for i = 0 to UBound(INCOME_ARRAY, 1)
					If INCOME_ARRAY(i).panel_name = "UNEA" Then
						If INCOME_ARRAY(i).income_type_code = "46" OR INCOME_ARRAY(i).income_type_code = "47" Then
							show_ti = FALSE
							for j = 0 to UBound(HH_MEMB_ARRAY, 1)
								If HH_MEMB_ARRAY(j).ref_number = INCOME_ARRAY(i).member_ref Then
									If first_ti = TRUE and memb_to_match = "" Then
										Text x_pos + 5, 331, 35, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number
										first_ti = FALSE
										show_ti = TRUE
									ElseIf memb_to_match = HH_MEMB_ARRAY(j).ref_number Then
										Text x_pos + 5, 331, 35, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number
										show_ti = TRUE
									Else
										PushButton x_pos, 330, 40, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number, HH_MEMB_ARRAY(j).button_one
									End If
									x_pos = x_pos + 40
								End If
							Next
							If show_ti = TRUE Then
								Text 110, 140, 160, 10, "HH Member: " & INCOME_ARRAY(i).member
								Text 280, 140, 105, 10, "Claim Number: " & INCOME_ARRAY(i).claim_number
								Text 110, 160, 120, 10, "Date Most Recent Income Received:"
								EditBox 235, 155, 45, 15, INCOME_ARRAY(i).most_recent_pay_date
								Text 290, 160, 130, 10, "How Much was the Most Recent Check:"
								EditBox 425, 155, 50, 15, INCOME_ARRAY(i).most_recent_pay_amt
								Text 110, 175, 110, 10, "Known Monthly Income: $" & INCOME_ARRAY(i).prosp_pay_total
								Text 110, 195, 75, 10, "Gross Monthly Income:"
								EditBox 190, 190, 35, 15, INCOME_ARRAY(i).pay_gross
								Text 235, 195, 70, 10, "Allowed Exclusions:"
								EditBox 305, 190, 35, 15, INCOME_ARRAY(i).expenses_allowed
								Text 355, 195, 85, 10, "Exclusions NOT Allowed:"
								EditBox 440, 190, 35, 15, INCOME_ARRAY(i).expenses_not_allowed
								Text 110, 210, 355, 10, "If the counted income is incorrect, press the Update Income button."
								Text 110, 225, 90, 10, "Start Date: " & INCOME_ARRAY(i).income_start_date
								Text 215, 225, 85, 10, "End Date: " & INCOME_ARRAY(i).income_end_date
								Text 110, 240, 50, 10, "Verification"
								Text 240, 240, 50, 10, "Verif Info"
								DropListBox 110, 250, 125, 45, unea_verif_droplist, INCOME_ARRAY(i).income_verification
								EditBox 240, 250, 235, 15, INCOME_ARRAY(i).verif_explaination
								Text 110, 265, 50, 10, "Income Notes"
								EditBox 110, 275, 365, 15, INCOME_ARRAY(i).income_notes
								Text 110, 295, 70, 10, "Review of Income"
								ComboBox 110, 305, 365, 45, "", INCOME_ARRAY(i).income_review

								PushButton 390, 135, 85, 15, "Update Income", update_ti_info_btn
							End If
						End If
					End If
				next
				If tribal_count = 0 Then
					Text 110, 140, 355, 20, "There are no Tribal Income panels known in MAXIS and no additional Tribal Income information has been added."
				End If
				PushButton 385, 347, 95, 13, "Add Tribal Income Information", add_another_unea_btn

				Text 44, 252, 43, 15, "Tribal - " & tribal_count
			End If
			If second_page_display = cs_unea Then								'=====================================================================================   UNEA - CHILD SUPPORT
				GroupBox 100, 125, 380, 220, "Child Support Income"

				x_pos = 110
				first_cs = TRUE
				for j = 0 to UBound(HH_MEMB_ARRAY, 1)
					show_cs = FALSE
					If HH_MEMB_ARRAY(j).clt_has_cs_income = TRUE Then
						If first_cs = TRUE and memb_to_match = "" Then
							first_cs = FALSE
							show_cs = TRUE
						ElseIf memb_to_match = HH_MEMB_ARRAY(j).ref_number Then
							show_cs = TRUE
						End If
					End If

					If show_cs = TRUE Then
						the_ref_to_use = HH_MEMB_ARRAY(j).ref_number

						Text 110, 140, 260, 10, "CS Income Received For " & HH_MEMB_ARRAY(j).full_name
						Text 110, 160, 30, 10, "Paid to:"
					    ComboBox 140, 155, 150, 45, memb_droplist, HH_MEMB_ARRAY(j).cs_paid_to
						Text 300, 160, 125, 10, "Is this Income Counted on this Case?"
					    DropListBox 425, 155, 50, 45, " "+chr(9)+"Yes"+chr(9)+"No", HH_MEMB_ARRAY(j).clt_cs_counted

						for i = 0 to UBound(INCOME_ARRAY, 1)
							If INCOME_ARRAY(i).member_ref = the_ref_to_use Then


								Select Case INCOME_ARRAY(i).income_type_code
									Case "08"
										Text 110, 175, 115, 10, "Known Direct Monthly Amt: $" & INCOME_ARRAY(i).prosp_pay_total
									Case "36"
										Text 110, 195, 130, 10, "Known Disbursed Monthly Amt: $" & INCOME_ARRAY(i).prosp_pay_total
										Text 250, 195, 50, 10, "Order Amount: "
										EditBox 300, 190, 40, 15, disb_order_amount
									Case "39"
										Text 110, 215, 120, 10, "Known Arrears Monthly Amt: $" & INCOME_ARRAY(i).prosp_pay_total
										Text 250, 215, 50, 10, "Order Amount: "
										EditBox 300, 210, 40, 15, arrears_order_amount
									Case "43"
									Case "45"
								End Select
							End If
						next
					End If
					If HH_MEMB_ARRAY(j).clt_has_cs_income = TRUE Then
						If show_cs = FALSE Then PushButton x_pos, 330, 40, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number, HH_MEMB_ARRAY(j).button_one
						If show_cs = TRUE Then Text x_pos + 5, 331, 35, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number
						x_pos = x_pos + 40
					End If
				next

				If cs_count = 0 Then
					Text 110, 140, 355, 20, "There are no Child Support panels known in MAXIS and no additional CS Income information has been added."
				End If
				PushButton 385, 347, 95, 13, "Add CS Income Information", add_another_unea_btn

				Text 32, 272, 55, 15, "Child Support - " & cs_count
			End If
			If second_page_display = ss_unea Then								'=====================================================================================   UNEA - SPOUSAL SUPPORT
				GroupBox 100, 125, 380, 220, "Spousal Support Income"

				x_pos = 110
				first_ss = TRUE
				for i = 0 to UBound(INCOME_ARRAY, 1)
					If INCOME_ARRAY(i).panel_name = "UNEA" Then
						If INCOME_ARRAY(i).income_type_code = "35" OR INCOME_ARRAY(i).income_type_code = "37" OR INCOME_ARRAY(i).income_type_code = "40" Then
							show_ss = FALSE
							for j = 0 to UBound(HH_MEMB_ARRAY, 1)
								If HH_MEMB_ARRAY(j).ref_number = INCOME_ARRAY(i).member_ref Then
									If first_ss = TRUE and memb_to_match = "" Then
										Text x_pos + 5, 331, 35, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number
										first_ss = FALSE
										show_ss = TRUE
									ElseIf memb_to_match = HH_MEMB_ARRAY(j).ref_number Then
										Text x_pos + 5, 331, 35, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number
										show_ss = TRUE
									Else
										PushButton x_pos, 330, 40, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number, HH_MEMB_ARRAY(j).button_one
									End If
									x_pos = x_pos + 40
								End If
							Next
							If show_ss = TRUE Then

							End If
						End If
					End If
				next
				If ss_count = 0 Then
					Text 110, 140, 355, 20, "There are no Spousal Support panels known in MAXIS and no additional Spousal Support Income information has been added."
				End If
				PushButton 385, 347, 95, 13, "Add Spousal Support Income Information", add_another_unea_btn

				Text 24, 292, 63, 15, "Spousal Support - " & ss_count
			End If
			If second_page_display = other_unea Then								'=====================================================================================   UNEA - TOTHER
				GroupBox 100, 125, 380, 220, "Other Unearned Income"

				x_pos = 110
				first_other = TRUE
				for i = 0 to UBound(INCOME_ARRAY, 1)
					If INCOME_ARRAY(i).panel_name = "UNEA" Then
						If INCOME_ARRAY(i).income_type_code = "06" OR INCOME_ARRAY(i).income_type_code = "18" OR INCOME_ARRAY(i).income_type_code = "19" OR INCOME_ARRAY(i).income_type_code = "20" OR INCOME_ARRAY(i).income_type_code = "21" OR INCOME_ARRAY(i).income_type_code = "22" OR INCOME_ARRAY(i).income_type_code = "23" OR INCOME_ARRAY(i).income_type_code = "24" OR INCOME_ARRAY(i).income_type_code = "25" OR INCOME_ARRAY(i).income_type_code = "26" OR INCOME_ARRAY(i).income_type_code = "27" OR INCOME_ARRAY(i).income_type_code = "28" OR INCOME_ARRAY(i).income_type_code = "29" OR INCOME_ARRAY(i).income_type_code = "30" OR INCOME_ARRAY(i).income_type_code = "31" OR INCOME_ARRAY(i).income_type_code = "44" OR INCOME_ARRAY(i).income_type_code = "48" OR INCOME_ARRAY(i).income_type_code = "49" Then
							show_other = FALSE
							for j = 0 to UBound(HH_MEMB_ARRAY, 1)
								If HH_MEMB_ARRAY(j).ref_number = INCOME_ARRAY(i).member_ref Then
									If first_other = TRUE and memb_to_match = "" Then
										Text x_pos + 5, 331, 35, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number
										first_other = FALSE
										show_other = TRUE
									ElseIf memb_to_match = HH_MEMB_ARRAY(j).ref_number Then
										Text x_pos + 5, 331, 35, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number
										show_other = TRUE
									Else
										PushButton x_pos, 330, 40, 10, "MEMB " & HH_MEMB_ARRAY(j).ref_number, HH_MEMB_ARRAY(j).button_one
									End If
									x_pos = x_pos + 40
								End If
							Next
							If show_other = TRUE Then
								Text 110, 140, 160, 10, "HH Member: " & INCOME_ARRAY(i).member
								Text 280, 140, 105, 10, "Claim Number: " & INCOME_ARRAY(i).claim_number
								Text 110, 160, 120, 10, "Date Most Recent Income Received:"
								EditBox 235, 155, 45, 15, INCOME_ARRAY(i).most_recent_pay_date
								Text 290, 160, 130, 10, "How Much was the Most Recent Check:"
								EditBox 425, 155, 50, 15, INCOME_ARRAY(i).most_recent_pay_amt
								Text 110, 175, 110, 10, "Known Monthly Income: $" & INCOME_ARRAY(i).prosp_pay_total
								Text 230, 175, 245, 10, "Othr Income Type: " & INCOME_ARRAY(i).income_type
								Text 110, 195, 75, 10, "Gross Monthly Income:"
								EditBox 190, 190, 35, 15, INCOME_ARRAY(i).pay_gross
								Text 235, 195, 70, 10, "Allowed Exclusions:"
								EditBox 305, 190, 35, 15, INCOME_ARRAY(i).expenses_allowed
								Text 355, 195, 85, 10, "Exclusions NOT Allowed:"
								EditBox 440, 190, 35, 15, INCOME_ARRAY(i).expenses_not_allowed
								Text 110, 210, 355, 10, "If the counted income is incorrect, press the Update Income button."
								Text 110, 225, 90, 10, "Start Date: " & INCOME_ARRAY(i).income_start_date
								Text 215, 225, 85, 10, "End Date: " & INCOME_ARRAY(i).income_end_date
								Text 110, 240, 50, 10, "Verification"
								Text 240, 240, 50, 10, "Verif Info"
								DropListBox 110, 250, 125, 45, unea_verif_droplist, INCOME_ARRAY(i).income_verification
								EditBox 240, 250, 235, 15, INCOME_ARRAY(i).verif_explaination
								Text 110, 265, 50, 10, "Income Notes"
								EditBox 110, 275, 365, 15, INCOME_ARRAY(i).income_notes
								Text 110, 295, 70, 10, "Review of Income"
								ComboBox 110, 305, 365, 45, "", INCOME_ARRAY(i).income_review

								PushButton 390, 135, 85, 15, "Update Income", update_other_info_btn
							End If
						End If
					End If
				next
				If other_UNEA_count = 0 Then
					Text 110, 140, 355, 20, "There are no Other UNEA panels known in MAXIS and no additional Other UNEA information has been added."
				End If
				PushButton 385, 347, 95, 13, "Add Other Income Information", add_another_unea_btn

				Text 44, 312, 43, 15, "Other - " & other_UNEA_count
			End If



		    Text 5, 350, 320, 10, "^^3 - For any type applied for or received - Click the button on the left to gather additional details."

			If second_page_display <> rsdi_unea Then PushButton 20, 130, 70, 15, "RSDI - " & rsdi_count, rsdi_btn
		    If second_page_display <> ssi_unea Then PushButton 20, 150, 70, 15, "SSI - " & ssi_count, ssi_btn
		    If second_page_display <> va_unea Then PushButton 20, 170, 70, 15, "VA - " & va_count, va_btn
		    If second_page_display <> ui_unea Then PushButton 20, 190, 70, 15, "UI - " & ui_count, ui_btn
		    If second_page_display <> wc_unea Then PushButton 20, 210, 70, 15, "WC - " & wc_count, wc_btn
		    If second_page_display <> ret_unea Then PushButton 20, 230, 70, 15, "Retirement - " & retirement_count, ret_btn
		    If second_page_display <> tribal_unea Then PushButton 20, 250, 70, 15, "Tribal - " & tribal_count, tribal_btn
		    If second_page_display <> cs_unea Then PushButton 20, 270, 70, 15, "Child Support - " & cs_count, cs_btn
		    If second_page_display <> ss_unea Then PushButton 20, 290, 70, 15, "Spousal Support - " & ss_count, ss_btn
		    If second_page_display <> other_unea Then PushButton 20, 310, 70, 15, "Other - " & other_UNEA_count, other_btn
		    If second_page_display <> main_unea Then PushButton 20, 330, 70, 15, "Main", main_btn




'DIALOG SAVING
' BeginDialog Dialog1, 0, 0, 550, 385, "Full Interview Questions"
'   Text 110, 135, 45, 10, "HH Member"
'   Text 275, 135, 40, 10, "Gross Amt"
'   Text 325, 135, 40, 10, "Exp Allowed"
'   Text 375, 135, 45, 10, "NOT Allowed"
'   Text 425, 135, 45, 10, "Counted Amt"
'   DropListBox 110, 145, 155, 45, "", member
'   EditBox 275, 145, 45, 15, gross_amt
'   EditBox 325, 145, 45, 15, exp_allowed
'   EditBox 375, 145, 45, 15, exp_not_allowed
'   EditBox 425, 145, 45, 15, counted_amt
'   Text 115, 160, 50, 10, "Verification"
'   Text 215, 160, 50, 10, "Verif Info"
'   Text 380, 160, 50, 10, "Claim Number"
'   DropListBox 115, 170, 95, 45, "", verif
'   EditBox 215, 170, 160, 15, verif_detail
'   EditBox 380, 170, 90, 15, claim_number
'   Text 115, 185, 50, 10, "Start Date"
'   Text 160, 185, 35, 10, "End Date"
'   Text 205, 185, 50, 10, "Pay Weekday"
'   Text 265, 185, 50, 10, "Income Notes"
'   EditBox 115, 195, 40, 15, Edit15
'   EditBox 160, 195, 40, 15, end_date
'   DropListBox 205, 195, 55, 45, "", pay_weekday
'   EditBox 265, 195, 205, 15, income_notes
'   Text 115, 210, 70, 10, "Review of Income"
'   ComboBox 115, 220, 355, 45, "", income_review
'   Text 115, 315, 70, 10, "Review of Income"
'   ComboBox 115, 325, 355, 45, "", Combo3
'   ButtonGroup ButtonPressed
'     PushButton 20, 330, 70, 15, "Main", main_btn
'   GroupBox 100, 125, 380, 220, "RSDI Income"
'   Text 5, 10, 195, 10, "^^1 - Enter the answers listed on the actual CAF from Q12"
'   ButtonGroup ButtonPressed
'     PushButton 20, 170, 70, 15, "VA", va_btn
'     PushButton 485, 10, 60, 15, "CAF Page 1", caf_page_one_btn
'     PushButton 485, 135, 60, 15, "CAF Page 1", Button27
'     PushButton 415, 365, 50, 15, "NEXT", next_btn
'     PushButton 465, 365, 80, 15, "Complete Interview", finish_interview_btn
'     PushButton 20, 290, 70, 15, "Spousal Support", ss_btn
'     PushButton 20, 310, 70, 15, "Other", other_btn
'     PushButton 385, 350, 95, 10, "Add RSDI Information", add_another_unea_btn
'   Text 5, 115, 35, 10, "^^2 - ASK - "
'   ButtonGroup ButtonPressed
'     PushButton 20, 190, 70, 15, "UI", ui_btn
'     PushButton 20, 270, 70, 15, "Child Support", cs_btn
'   Text 5, 350, 320, 10, "^^3 - For any type applied for or received - Click the button on the left to gather additional details."
'   Text 40, 115, 280, 10, "'Has anyone in the household applied for or receive any ..."
'   ButtonGroup ButtonPressed
'     PushButton 20, 210, 70, 15, "WC", wc_btn
'     PushButton 20, 130, 70, 15, "RSDI", rsdi_btn
'     PushButton 20, 230, 70, 15, "Retirement", ret_btn
'     PushButton 20, 150, 70, 15, "SSI", ssi_btn
'     PushButton 20, 250, 70, 15, "Tribal", tribal_btn
' EndDialog

' Text 5, 10, 305, 10, "^^1 - Enter the answers listed on the actual CAF form for Q5 into the 'Answer on the CAF' field."
' Text 20, 25, 335, 20, "Q. 12. Has anyone in the household applied for or does anyone get any of the following types of income each month?"
' Text 370, 30, 65, 10, "Answer on the CAF"
' DropListBox 435, 25, 40, 45, "Yes"+chr(9)+"No"+chr(9)+"Blank", caf_answer
' Text 5, 50, 35, 10, "^^2 - ASK - "
' Text 40, 50, 280, 10, "'Is anyone in the household disabled or have a physical or mental health condition?"
' Text 40, 70, 70, 10, "Confirm CAF Answer"
' ComboBox 110, 65, 365, 45, "", confirm_caf_answer

		End If
		If page_display = show_q_13 Then										'THIS IS THE INFORMATION FOR PAGE  ----------------------------------------------------------------------------------------------------------------------------------
			Text 507, 207, 60, 13, "Q. 13"

			Text 5, 10, 305, 10, "^^1 - Enter the answers listed on the actual CAF form for Q13 into the 'Answer on the CAF' field."
			Text 20, 25, 335, 20, "Q. 13. Does anyone in the household have or expect to get any loans, scholarships or grants for attending school?"
			Text 370, 30, 65, 10, "Answer on the CAF"
			DropListBox 435, 25, 40, 45, caf_answer_droplist, q13_caf_answer
			Text 5, 50, 35, 10, "^^2 - ASK - "
			Text 40, 50, 280, 10, "'Is anyone in the household disabled or have a physical or mental health condition?"
			Text 40, 70, 70, 10, "Confirm CAF Answer"
			ComboBox 110, 65, 365, 45, "", q13_confirm_caf_answer


		    Text 5, 95, 345, 10, "^^3 - ENTER information about student income using the 'Details - Add' button or 'Details - Update' button."

			y_pos = 115
			For i = 0 to UBound(HH_MEMB_ARRAY, 1)
				If HH_MEMB_ARRAY(i).stin_exists = FALSE and HH_MEMB_ARRAY(i).stec_exists = FALSE Then
					Text 20, y_pos, 340, 10, "MEMB " & HH_MEMB_ARRAY(i).ref_number & " - " & HH_MEMB_ARRAY(i).full_name & " - No STIN or STEC"
				ElseIf HH_MEMB_ARRAY(i).stin_exists = TRUE and HH_MEMB_ARRAY(i).stec_exists = FALSE Then
					Text 20, y_pos, 340, 10, "MEMB " & HH_MEMB_ARRAY(i).ref_number & " - " & HH_MEMB_ARRAY(i).full_name & " - STIN: $" & HH_MEMB_ARRAY(i).total_stin & " - No STEC"
				ElseIf HH_MEMB_ARRAY(i).stin_exists = FALSE and HH_MEMB_ARRAY(i).stec_exists = TRUE Then
					Text 20, y_pos, 340, 10, "MEMB " & HH_MEMB_ARRAY(i).ref_number & " - " & HH_MEMB_ARRAY(i).full_name & " - No STIN - STEC: $" & HH_MEMB_ARRAY(i).total_stec
				ElseIf HH_MEMB_ARRAY(i).stin_exists = TRUE and HH_MEMB_ARRAY(i).stec_exists = TRUE Then
					Text 20, y_pos, 340, 10, "MEMB " & HH_MEMB_ARRAY(i).ref_number & " - " & HH_MEMB_ARRAY(i).full_name & " - STIN: $" & HH_MEMB_ARRAY(i).total_stin & " - STEC: $" & HH_MEMB_ARRAY(i).total_stec
				End If
		    	If HH_MEMB_ARRAY(i).stin_exists = TRUE OR HH_MEMB_ARRAY(i).stec_exists = TRUE Then
					PushButton 370, y_pos-2, 105, 13, "Details - Update", HH_MEMB_ARRAY(i).button_one
				Else
					PushButton 370, y_pos-2, 105, 13, "Details - Add", HH_MEMB_ARRAY(i).button_one
				End If
				y_pos = y_pos + 15
			Next
			y_pos = y_pos + 5

		    Text 20, y_pos, 265, 10, "Any HH Member that has known STIN or STEC must have details updated."

		End If
		If page_display = show_q_14_15 Then										'THIS IS THE INFORMATION FOR PAGE  ----------------------------------------------------------------------------------------------------------------------------------
			Text 495, 222, 60, 13, "Q. 14 and 15"

			Text 5, 10, 305, 10, "^^1 - Enter the answers listed on the actual CAF form for Q14 into the 'Answer on the CAF' field."
		    Text 20, 20, 225, 10, "Q. 14. Does your household have the following housing expenses?"

		    Text 5, 30, 255, 10, "^^2 - If any are 'YES' then ASK the amount and ENTER the amount answered."
		    Text 25, 50, 125, 10, "Rent (include mobild home lot rental)"
		    DropListBox 155, 45, 40, 45, caf_answer_droplist, q14_rent_caf_answer
		    EditBox 205, 45, 35, 15, q14_rent_caf_response
		    Text 25, 65, 125, 10, "Mortgage/Contract for Deed Payment"
		    DropListBox 155, 60, 40, 45, caf_answer_droplist, q14_mort_caf_answer
		    EditBox 205, 60, 35, 15, q14_mort_caf_response
		    Text 25, 80, 125, 10, "Homeowner's Insurance"
		    DropListBox 155, 75, 40, 45, caf_answer_droplist, q14_ins_caf_answer
		    EditBox 205, 75, 35, 15, q14_ins_caf_response
		    Text 25, 95, 125, 10, "Real Estate Taxes"
		    DropListBox 155, 90, 40, 45, caf_answer_droplist, q14_tax_caf_answer
		    EditBox 205, 90, 35, 15, q14_tax_caf_response

		    Text 255, 50, 105, 10, "Rental or Secontion 8 Subsidy"
		    DropListBox 365, 45, 40, 45, caf_answer_droplist, q14_subs_caf_answer
		    EditBox 415, 45, 35, 15, q14_subs_caf_response
		    Text 255, 65, 100, 10, "Association Fees"
		    DropListBox 365, 60, 40, 45, caf_answer_droplist, q14_fees_caf_answer
		    EditBox 415, 60, 35, 15, q14_fees_caf_response
		    Text 255, 80, 95, 10, "Room and/or Board"
		    DropListBox 365, 75, 40, 45, caf_answer_droplist, q14_room_caf_answer
		    EditBox 415, 75, 35, 15, q14_room_caf_response
			Text 255, 95, 105, 20, "CONFIM - Do you get help paying rent?"
			DropListBox 365, 95, 40, 45, caf_answer_droplist, q14_confirm_subsidy
			EditBox 415, 95, 35, 15, q14_confirm_subsidy_amount

		    Text 5, 115, 455, 10, "^^3 - ASK - 'Explain how you pay your housing expenses.' and REVIEW Shelter Expenses entered or in MAXIS."
		    Text 20, 135, 70, 10, "Housing Explanation"
		    ComboBox 95, 130, 380, 45, "", q14_confirm_caf_answer
			y_pos = 165
			grp_len = 15
			for i = 0 to UBound(HH_MEMB_ARRAY, 1)
				If HH_MEMB_ARRAY(i).shel_exists = TRUE Then
					Text 30, y_pos, 440, 10, "MEMB " & HH_MEMB_ARRAY(i).ref_number & " - " & HH_MEMB_ARRAY(i).full_name & ": " & HH_MEMB_ARRAY(i).shel_summary
					y_pos = y_pos + 15
					grp_len = grp_len + 15
				End If
			Next
			If grp_len = 15 Then
				Text 30, y_pos, 440, 10, "NO SHELTER EXPENSES LISTED IN MAXIS OR ADDED TO THE SCRIPT."
				y_pos = y_pos + 15
				grp_len = grp_len + 15
			End If
			y_pos = y_pos + 5
			GroupBox 20, 150, 455, grp_len, "Already Known Shelter Expenses - Added or listed in MAXIS"
		    ' Text 30, 165, 440, 10, "MEMB 01 - CLIENT FULL NAME HERE - Amount: $400"
		    ' Text 30, 180, 440, 10, "MEMB 01 - CLIENT FULL NAME HERE - Amount: $400"
		    PushButton 350, y_pos, 125, 10, "Update Shelter Expense Information", update_shel_btn
			y_pos = y_pos + 15
			Text 5, y_pos, 310, 10, "^^4 - Enter the answers listed on the actual CAF form for Q15 into the 'Answer on the CAF' field."
		    Text 20, y_pos + 10, 295, 10, "Q. 15. Does your household have the following utility expenses any time during the year?"
			y_pos = y_pos + 30
		    Text 20, y_pos, 85, 10, "Heating/Air Conditioning"
		    DropListBox 110, y_pos - 5, 40, 45, caf_answer_droplist, q15_h_ac_caf_answer
			Text 180, y_pos, 85, 10, "Electricity"
			DropListBox 270, y_pos - 5, 40, 45, caf_answer_droplist, q15_e_caf_answer
			Text 345, y_pos, 85, 10, "Cooking Fuel"
			DropListBox 435, y_pos - 5, 40, 45, caf_answer_droplist, q15_cf_caf_answer
			y_pos = y_pos + 15
		    Text 20, y_pos, 85, 10, "Water and Sewer"
		    DropListBox 110, y_pos - 5, 40, 45, caf_answer_droplist, q15_ws_caf_answer
		    Text 180, y_pos, 85, 10, "Garbage Removal"
		    DropListBox 270, y_pos - 5, 40, 45, caf_answer_droplist, q15_gr_caf_answer
		    Text 345, y_pos, 85, 10, "Phone/Cell Phone"
		    DropListBox 435, y_pos - 5, 40, 45, caf_answer_droplist, q15_p_caf_answer
			y_pos = y_pos + 15
		    Text 75, y_pos, 355, 10, "Did anyone in the household receive Energy Assistance (LIHEAP) of more than $20 in the past 12 months?"
		    DropListBox 435, y_pos - 5, 40, 45, caf_answer_droplist, q15_liheap_caf_answer
			y_pos = y_pos + 15

			Text 5, y_pos, 270, 10, "^^5 - ASK - 'Does anyone in the household pay ...'  RECORD the verbal responses"
			y_pos = y_pos + 20
		    Text 20, y_pos, 85, 10, "Heating"
		    DropListBox 110, y_pos - 5, 40, 45, caf_answer_droplist, q15_h_caf_response
			Text 180, y_pos, 85, 10, "Electricity"
			DropListBox 270, y_pos - 5, 40, 45, caf_answer_droplist, q15_e_caf_response
			Text 345, y_pos, 85, 10, "Cooking Fuel"
			DropListBox 435, y_pos - 5, 40, 45, caf_answer_droplist, q15_cf_caf_response
			y_pos = y_pos + 15
		    Text 20, y_pos, 85, 10, "Air Conditioning"
		    DropListBox 110, y_pos - 5, 40, 45, caf_answer_droplist, q15_ac_caf_response
		    Text 180, y_pos, 85, 10, "Garbage Removal"
		    DropListBox 270, y_pos - 5, 40, 45, caf_answer_droplist, q15_gr_caf_response
		    Text 345, y_pos, 85, 10, "Phone/Cell Phone"
		    DropListBox 435, y_pos - 5, 40, 45, caf_answer_droplist, q15_p_caf_response
			y_pos = y_pos + 15
			Text 20, y_pos, 85, 10, "Water and Sewer"
			DropListBox 110, y_pos - 5, 40, 45, caf_answer_droplist, q15_ws_caf_response
		    Text 170, y_pos + 5, 265, 10, "Did your household receive any help in paying for your energy or power bills?"
		    DropListBox 435, y_pos, 40, 45, caf_answer_droplist, q15_liheap_caf_response
			y_pos = y_pos + 15
		    PushButton 20, y_pos, 130, 10, "Utilities are Complicated", utility_detail_btn

		End If
		If page_display = show_q_16_18 Then										'THIS IS THE INFORMATION FOR PAGE  ----------------------------------------------------------------------------------------------------------------------------------
			Text 487, 237, 60, 13, "Q. 16, 17, and 18"

			Text 5, 10, 305, 10, "^^1 - Enter the answers listed on the actual CAF form for Q16 into the 'Answer on the CAF' field."
		    Text 365, 10, 65, 10, "Answer on the CAF"
		    DropListBox 435, 5, 40, 45, caf_answer_droplist, q16_caf_answer
		    Text 20, 20, 445, 20, "Q. 16. Do you or anyone living with you have costs for care of a child(ren) because you or they are working, looking for work or going to school? The Child Care Assistance Program (CCAP) may help pay child care costs."
		    Text 5, 45, 35, 10, "^^2 - ASK - "
		    Text 40, 45, 210, 10, "'Does anyone in your household have costs for care of a child?'"
		    DropListBox 260, 40, 40, 45, caf_answer_droplist, q16_caf_confirm
		    Text 40, 60, 100, 10, "Additional detail from answer:"
		    EditBox 140, 55, 335, 15, q16_caf_confirm_notes
		    Text 5, 80, 305, 10, "^^3 - Enter the answers listed on the actual CAF form for Q17 into the 'Answer on the CAF' field."
		    Text 370, 80, 65, 10, "Answer on the CAF"
		    DropListBox 435, 75, 40, 45, caf_answer_droplist, q17_caf_answer
		    Text 20, 90, 460, 20, "Q. 17. Do you or anyone living with you have costs for care of an ill or disabled adult because you are working, looking for work or going to school?"
		    Text 5, 115, 35, 10, "^^4 - ASK - "
		    Text 40, 115, 215, 10, "'Does anyone in your household have cost for care of an adult?'"
		    DropListBox 255, 110, 40, 45, caf_answer_droplist, q17_caf_confirm
		    Text 40, 130, 100, 10, "Additional detail from answer:"
		    EditBox 140, 125, 335, 15, q17_caf_confirm_notes
		    Text 5, 145, 255, 10, "^^5 - REVIEW Known Information (listed here) and UPDATE based on answer"
		    GroupBox 20, 155, 455, 55, "DCEX in MAXIS or added manually"
		    Text 30, 170, 435, 10, "Care Cost for MEMB 03 - $500 per month"
		    Text 30, 180, 435, 10, "Care Cost for MEMB 03 - $500 per month"
		    Text 25, 200, 100, 10, "Confirm Information is Correct:"
		    DropListBox 130, 195, 205, 45, ""+chr(9)+""+chr(9)+""+chr(9)+"", List7
		    PushButton 345, 195, 120, 10, "Update Child Care Information", update_DCEX_info_btn
		    Text 5, 220, 305, 10, "^^6 - Enter the answers listed on the actual CAF form for Q18 into the 'Answer on the CAF' field."
		    Text 20, 230, 320, 20, "Q. 18. Does anyone in the household pay court-ordered child support, spousal support, child care support, medical support or contribute to a tax-dependent who does not live in your home?"
		    Text 370, 230, 65, 10, "Answer on the CAF"
		    DropListBox 435, 225, 40, 45, caf_answer_droplist, q18_caf_answer
		    Text 5, 255, 35, 10, "^^7 - ASK - "
		    Text 40, 255, 315, 10, "'Does anyone in your household pay court ordered expenses for someone outside of the home?'"
		    DropListBox 360, 250, 40, 45, caf_answer_droplist, List6
		    Text 40, 270, 100, 10, "Additional detail from answer:"
		    EditBox 140, 265, 335, 15, Edit3
		    Text 5, 285, 255, 10, "^^8 - REVIEW Known Information (listed here) and UPDATE based on answer"
		    GroupBox 20, 295, 455, 55, "COEX in MAXIS or added manually"
		    Text 30, 310, 435, 10, "Care Cost for MEMB 03 - $500 per month"
		    Text 30, 320, 435, 10, "Care Cost for MEMB 03 - $500 per month"
		    Text 25, 340, 100, 10, "Confirm Information is Correct:"
		    DropListBox 130, 335, 205, 45, "", List8
		    PushButton 345, 335, 120, 10, "Update Child Care Information", Button7

			'
			' Text 5, 10, 305, 10, "^^1 - Enter the answers listed on the actual CAF form for Q5 into the 'Answer on the CAF' field."
			' Text 20, 25, 335, 20, "Q. 16. Do you or anyone living with you have costs for care of a child(ren) because you or they are working, looking for work or going to school? The Child Care Assistance Program (CCAP) may help pay child care costs."
			' Text 370, 30, 65, 10, "Answer on the CAF"
			' DropListBox 435, 25, 40, 45, caf_answer_droplist, q16_caf_answer
			' Text 5, 50, 35, 10, "^^2 - ASK - "
			' Text 40, 50, 280, 10, "'Is anyone in the household disabled or have a physical or mental health condition?"
			' Text 40, 70, 70, 10, "Confirm CAF Answer"
			' ComboBox 110, 65, 365, 45, "", q16_confirm_caf_answer
			'
			'
			'
			' Text 5, 90, 305, 10, "^^1 - Enter the answers listed on the actual CAF form for Q5 into the 'Answer on the CAF' field."
			' Text 20, 105, 335, 20, "Q. 17. Do you or anyone living with you have costs for care of an ill or disabled adult because you are working, looking for work or going to school?"
			' Text 370, 110, 65, 10, "Answer on the CAF"
			' DropListBox 435, 105, 40, 45, caf_answer_droplist, q17_caf_answer
			' Text 5, 130, 35, 10, "^^2 - ASK - "
			' Text 40, 130, 280, 10, "'Is anyone in the household disabled or have a physical or mental health condition?"
			' Text 40, 150, 70, 10, "Confirm CAF Answer"
			' ComboBox 110, 145, 365, 45, "", q17_confirm_caf_answer
			'
			'
			'
			' Text 5, 170, 305, 10, "^^1 - Enter the answers listed on the actual CAF form for Q5 into the 'Answer on the CAF' field."
			' Text 20, 185, 335, 20, "Q. 18. Does anyone in the household pay court-ordered child support, spousal support, child care support, medical support or contribute to a tax-dependent who does not live in your home?"
			' Text 370, 190, 65, 10, "Answer on the CAF"
			' DropListBox 435, 185, 40, 45, caf_answer_droplist, q18_caf_answer
			' Text 5, 210, 35, 10, "^^2 - ASK - "
			' Text 40, 210, 280, 10, "'Is anyone in the household disabled or have a physical or mental health condition?"
			' Text 40, 230, 70, 10, "Confirm CAF Answer"
			' ComboBox 110, 225, 365, 45, "", q18_confirm_caf_answer




' BeginDialog Dialog1, 0, 0, 550, 385, "Full Interview Questions"
'   ButtonGroup ButtonPressed
'     PushButton 415, 365, 50, 15, "NEXT", next_btn
'     PushButton 465, 365, 80, 15, "Complete Interview", finish_interview_btn
'     PushButton 485, 10, 60, 15, "CAF Page 1", caf_page_one_btn
'     PushButton 485, 135, 60, 15, "CAF Page 1", Button27
'   Text 5, 10, 305, 10, "^^1 - Enter the answers listed on the actual CAF form for Q16 into the 'Answer on the CAF' field."
'   Text 365, 10, 65, 10, "Answer on the CAF"
'   DropListBox 435, 5, 40, 45, "caf_answer_droplist", q16_caf_answer
'   Text 20, 20, 445, 20, "Q. 16. Do you or anyone living with you have costs for care of a child(ren) because you or they are working, looking for work or going to school? The Child Care Assistance Program (CCAP) may help pay child care costs."
'   Text 5, 45, 35, 10, "^^2 - ASK - "
'   Text 40, 45, 210, 10, "'Does anyone in your household have costs for care of a child?'"
'   DropListBox 260, 40, 40, 45, "caf_answer_droplist", List4
'   Text 40, 60, 100, 10, "Additional detail from answer:"
'   EditBox 140, 55, 335, 15, Edit1
'   Text 5, 80, 305, 10, "^^3 - Enter the answers listed on the actual CAF form for Q17 into the 'Answer on the CAF' field."
'   Text 370, 80, 65, 10, "Answer on the CAF"
'   DropListBox 435, 75, 40, 45, "caf_answer_droplist", q17_caf_answer
'   Text 20, 90, 460, 20, "Q. 17. Do you or anyone living with you have costs for care of an ill or disabled adult because you are working, looking for work or going to school?"
'   Text 5, 115, 35, 10, "^^4 - ASK - "
'   Text 40, 115, 215, 10, "'Does anyone in your household have cost for care of an adult?'"
'   DropListBox 255, 110, 40, 45, "caf_answer_droplist", List5
'   Text 40, 130, 100, 10, "Additional detail from answer:"
'   EditBox 140, 125, 335, 15, Edit2
'   Text 5, 145, 255, 10, "^^5 - REVIEW Known Information (listed here) and UPDATE based on answer"
'   GroupBox 20, 155, 455, 55, "DCEX in MAXIS or added manually"
'   Text 30, 170, 435, 10, "Care Cost for MEMB 03 - $500 per month"
'   Text 30, 180, 435, 10, "Care Cost for MEMB 03 - $500 per month"
'   Text 25, 200, 100, 10, "Confirm Information is Correct:"
'   DropListBox 130, 195, 205, 45, "", List7
'   ButtonGroup ButtonPressed
'     PushButton 345, 195, 120, 10, "Update Child Care Information", update_DCEX_info_btn
'   Text 5, 220, 305, 10, "^^6 - Enter the answers listed on the actual CAF form for Q18 into the 'Answer on the CAF' field."
'   Text 20, 230, 320, 20, "Q. 18. Does anyone in the household pay court-ordered child support, spousal support, child care support, medical support or contribute to a tax-dependent who does not live in your home?"
'   Text 370, 230, 65, 10, "Answer on the CAF"
'   DropListBox 435, 225, 40, 45, "caf_answer_droplist", q18_caf_answer
'   Text 5, 255, 35, 10, "^^7 - ASK - "
'   Text 40, 255, 315, 10, "'Does anyone in your household pay court ordered expenses for someone outside of the home?'"
'   DropListBox 360, 250, 40, 45, "caf_answer_droplist", List6
'   Text 40, 270, 100, 10, "Additional detail from answer:"
'   EditBox 140, 265, 335, 15, Edit3
'   Text 5, 285, 255, 10, "^^8 - REVIEW Known Information (listed here) and UPDATE based on answer"
'   GroupBox 20, 295, 455, 55, "COEX in MAXIS or added manually"
'   Text 30, 310, 435, 10, "Care Cost for MEMB 03 - $500 per month"
'   Text 30, 320, 435, 10, "Care Cost for MEMB 03 - $500 per month"
'   Text 25, 340, 100, 10, "Confirm Information is Correct:"
'   DropListBox 130, 335, 205, 45, "", List8
'   ButtonGroup ButtonPressed
'     PushButton 345, 335, 120, 10, "Update Child Care Information", Button7
' EndDialog


		End If
		If page_display = show_q_19 Then										'THIS IS THE INFORMATION FOR PAGE  ----------------------------------------------------------------------------------------------------------------------------------
			Text 507, 252, 60, 13, "Q. 19"

			Text 5, 10, 305, 10, "^^1 - Enter the answers listed on the actual CAF form for Q5 into the 'Answer on the CAF' field."
			Text 20, 25, 335, 20, "Q. 19. For SNAP only: Does anyone in the household have medical expenses?"
			Text 370, 30, 65, 10, "Answer on the CAF"
			DropListBox 435, 25, 40, 45, caf_answer_droplist, q19_caf_answer
			Text 5, 50, 35, 10, "^^2 - ASK - "
			Text 40, 50, 280, 10, "'Does anyone in the household have medical expenses?'"
			Text 40, 70, 70, 10, "Confirm CAF Answer"
			ComboBox 110, 65, 365, 45, "", q19_confirm_caf_answer

			Text 5, 90, 315, 10, "^^3 - CONFIRM - Review details of known STWK or add details if not currently known."

			grp_len = 180

			If case_has_fmed_panel = FALSE Then
				GroupBox 50, 105, 425, 24, "FMED on CASE " & MAXIS_case_number
				Text 75, 115, 250, 10, "NO FMED PANELS OR OTHER FMED DETAILS ON THIS CASE"
			Else
				GroupBox 50, 105, 425, grp_len, "FMED for " & HH_MEMB_ARRAY(fmed_selected).full_name

				Text 55, 120, 50, 10, "Type"
				Text 175, 120, 60, 10, "Verif"
				Text 250, 120, 80, 10, "Category"
				Text 325, 120, 120, 10, "Expense Amount"
				y_pos = 140

				For clts_fmed = 0 to UBound(HH_MEMB_ARRAY(fmed_selected).fmed_type)
					Text 55, y_pos, 50, 10, HH_MEMB_ARRAY(fmed_selected).fmed_type(clts_fmed)
					Text 175, y_pos, 60, 10, HH_MEMB_ARRAY(fmed_selected).fmed_verif(clts_fmed)
					Text 250, y_pos, 80, 10, HH_MEMB_ARRAY(fmed_selected).fmed_catgry(clts_fmed)
					Text 325, y_pos, 120, 10, HH_MEMB_ARRAY(fmed_selected).fmed_expense(clts_fmed)
					y_pos = y_pos + 15
				Next

				PushButton 55, 270, 125, 10, "Update FMED Information", fmed_info_change
				PushButton 410, 270, 60, 10, "NEXT FMED", next_fmed_btn

				btn_pos = 110
				For i = 0 to Ubound(HH_MEMB_ARRAY, 1)
					If i = fmed_selected Then
						Text 9, btn_pos+1, 40, 10, "MEMB " & HH_MEMB_ARRAY(i).ref_number
					ElseIf HH_MEMB_ARRAY(i).fmed_exists = TRUE Then
						PushButton 5, btn_pos, 40, 10, "MEMB " & HH_MEMB_ARRAY(i).ref_number, HH_MEMB_ARRAY(i).button_one
					End If
					btn_pos = btn_pos + 10
				Next
				btn_pos = btn_pos + 20
			End If

			PushButton 5, 290, 150, 10, "ADD STWK FOR ANOTHER MEMBER", add_new_stwk_btn

		End If

		If page_display = show_q_20_21 Then										'THIS IS THE INFORMATION FOR PAGE  ----------------------------------------------------------------------------------------------------------------------------------
			Text 495, 267, 60, 15, "Q. 20 and 21"
		    Text 5, 10, 200, 10, "^^1 - Enter the answers listed on the actual CAF form for Q20"
		    Text 20, 20, 335, 10, "Q. 20. Does anyone in the household own, or is anyone buying, any of the following?"
		    Text 30, 40, 20, 10, "Cash"
		    DropListBox 70, 35, 40, 45, caf_answer_droplist, cash_caf_answer
		    Text 250, 40, 55, 10, "Bank accounts"
			Text 350, 40, 170, 10, "(savings, checking, debit card, etc.)"
		    DropListBox 305, 35, 40, 45, caf_answer_droplist, acct_caf_answer

		    Text 30, 55, 40, 10, "Securities"
			Text 115, 55, 170, 10, "(stocks, bonds, annuities, 401K, etc.)"
		    DropListBox 70, 50, 40, 45, caf_answer_droplist, secu_caf_answer
		    Text 250, 55, 35, 10, "Vehicles"
			Text 350, 52, 130, 20, "(cars, trucks, motorcycles, campers, trailers, etc.)"
		    DropListBox 305, 50, 40, 45, caf_answer_droplist, cars_caf_answer
		    Text 5, 75, 35, 10, "^^2 - ASK - "
		    Text 40, 75, 250, 10, "'Does anyone own or is buying any of the following?' and CONFIRM details."

			cash_count = 0
			acct_count = 0
			secu_count = 0
			cars_count = 0
			' MsgBox "UBound - " & UBound(ASSET_ARRAY, 1)
			for i = 0 to UBound(ASSET_ARRAY)
				If ASSET_ARRAY(i).panel_name = "CASH" Then cash_count = cash_count + 1
				If ASSET_ARRAY(i).panel_name = "ACCT" Then acct_count = acct_count + 1
				' If ASSET_ARRAY(i).panel_name = "ACCT" Then MsgBox "ACCT" & vbNewLine & i
				If ASSET_ARRAY(i).panel_name = "SECU" Then secu_count = secu_count + 1
				If ASSET_ARRAY(i).panel_name = "CARS" Then cars_count = cars_count + 1
				' If ASSET_ARRAY(i).panel_name = "ACCT" Then acct_count = acct_count + 1
			next
			If second_page_display = main_asset Then								'=====================================================================================   Account - CASH
			    Text 70, 107, 20, 10, "Cash"
			    ComboBox 130, 105, 290, 45, "", cash_confirm_response
			    Text 70, 122, 55, 10, "Bank accounts"
			    ComboBox 130, 120, 290, 45, "", acct_confirm_response
			    Text 70, 137, 40, 10, "Securities"
			    ComboBox 130, 135, 290, 45, "", secu_confirm_response
			    Text 70, 152, 35, 10, "Vehicles"
			    ComboBox 130, 150, 290, 45, "", cars_confirm_response
				Text 70, 167, 55, 10, "Real Estate"
			    ComboBox 130, 165, 290, 45, "", rest_confirm_response

				Text 70, 185, 350, 20, "Enter the details of the verbal answer here. Capture any details about each specific asset usng the buttons to the left."

				Text 30, 92, 40, 10, "MAIN"
			End If
			If second_page_display = cash_asset Then								'=====================================================================================   Account - CASH
				GroupBox 70, 90, 405, 190, "Cash"

				x_pos = 75
				first_cash = TRUE
				for i = 0 to UBound(ASSET_ARRAY)
					If ASSET_ARRAY(i).panel_name = "CASH" Then
						show_cash = FALSE
						If memb_to_match = "" and first_cash = TRUE Then
							Text x_pos + 5, 276, 40, 10, "CASH " & ASSET_ARRAY(i).member_ref
							first_cash = FALSE
							show_cash = TRUE
						ElseIf ASSET_ARRAY(i).member_ref = memb_to_match  Then
							Text x_pos + 5, 276, 40, 10, "CASH " & ASSET_ARRAY(i).member_ref
							show_cash = TRUE
						Else
							PushButton x_pos, 275, 40, 10, "CASH " & ASSET_ARRAY(i).member_ref, ASSET_ARRAY(i).asset_btn_one
						End If
						' MsgBox "Panel: " & vbCr & "Member - " & ASSET_ARRAY(i).member_ref & vbCr & "Instance - " & ASSET_ARRAY(i).panel_instance & vbCr & "Member to match - " & memb_to_match & vbCr & "Inst to match - " & inst_to_match
						x_pos = x_pos + 40

						If show_cash = TRUE Then
							Text 80, 105, 160, 10, "HH Member: " & ASSET_ARRAY(i).member

							Text 80, 125, 160, 10, "Cash Amount: " & ASSET_ARRAY(i).cash_value

							PushButton 360, 105, 105, 13, "Update this CASH Panel", update_asset_btn
						End If

					End If
				next
				If cash_count = 0 Then
					Text 110, 105, 300, 20, "There are no CASH panels known in MAXIS and no additional CASH information has been added."
				End If
				PushButton 360, 87, 105, 13, "Add New Cash Information", add_another_asset_btn
				Text 25, 107, 40, 10, "CASH - " & cash_count
			End If
			If second_page_display = acct_asset Then								'=====================================================================================   Account - ACCT
				GroupBox 70, 90, 405, 190, "Bank Account"

				x_pos = 75
				first_acct = TRUE
				for i = 0 to UBound(ASSET_ARRAY)
					If ASSET_ARRAY(i).panel_name = "ACCT" Then
						show_acct = FALSE
						If memb_to_match = "" and inst_to_match = "" AND first_acct = TRUE Then
							Text x_pos + 5, 276, 45, 10, "ACCT " & ASSET_ARRAY(i).member_ref & " " & ASSET_ARRAY(i).panel_instance
							first_acct = FALSE
							show_acct = TRUE
						ElseIf ASSET_ARRAY(i).member_ref = memb_to_match AND ASSET_ARRAY(i).panel_instance = inst_to_match Then
							Text x_pos + 5, 276, 45, 10, "ACCT " & ASSET_ARRAY(i).member_ref & " " & ASSET_ARRAY(i).panel_instance
							show_acct = TRUE
						Else
							PushButton x_pos, 275, 45, 10, "ACCT " & ASSET_ARRAY(i).member_ref & " " & ASSET_ARRAY(i).panel_instance, ASSET_ARRAY(i).asset_btn_one
						End If
						' MsgBox "Panel: " & vbCr & "Member - " & ASSET_ARRAY(i).member_ref & vbCr & "Instance - " & ASSET_ARRAY(i).panel_instance & vbCr & "Member to match - " & memb_to_match & vbCr & "Inst to match - " & inst_to_match
						x_pos = x_pos + 45

						If show_acct = TRUE Then
							Text 80, 105, 160, 10, "HH Member: " & ASSET_ARRAY(i).member

							Text 80, 125, 160, 10, "Account Type: " & ASSET_ARRAY(i).asset_type
							Text 80, 140, 160, 10, "Account Number: " & ASSET_ARRAY(i).account_number
							Text 80, 155, 160, 10, "Account Location: " & ASSET_ARRAY(i).acct_location
							Text 80, 175, 160, 10, "Balance: $ " & ASSET_ARRAY(i).acct_balance
							Text 250, 175, 160, 10, "Ver: " & ASSET_ARRAY(i).asset_verification

							Text 80, 190, 160, 10, "As Of: " & ASSET_ARRAY(i).acct_balance_date
							Text 80, 205, 120, 10, "Withdrawal Penalty: " & ASSET_ARRAY(i).withdraw_penalty
							Text 210, 205, 75, 10, "(Y/N): " & ASSET_ARRAY(i).withdraw_yn
							Text 285, 205, 160, 10, "Ver: " & ASSET_ARRAY(i).withdraw_verif

							Text 80, 225, 160, 10, "Count (Y/N): Cash: " & ASSET_ARRAY(i).count_cash_yn
							Text 240, 225, 50, 10, "FS: " & ASSET_ARRAY(i).count_snap_yn
							Text 290, 225, 50, 10, "HC: " & ASSET_ARRAY(i).count_hc_yn
							Text 340, 225, 50, 10, "GRH: " & ASSET_ARRAY(i).count_grh_yn
							Text 390, 225, 50, 10, "IV-E: " & ASSET_ARRAY(i).count_ive_yn

							Text 80, 240, 160, 10, "Joint Owner (Y/N): " & ASSET_ARRAY(i).joint_owners_yn
							Text 250, 240, 160, 10, "Share Ratio: " & ASSET_ARRAY(i).share_ratio

							Text 80, 260, 160, 10, "Next Interest Date (MM/YY): " & ASSET_ARRAY(i).next_interest_date

							PushButton 360, 105, 105, 13, "Update this ACCT Panel", update_asset_btn
						End If
					End If
				next
				If acct_count = 0 Then
					Text 110, 105, 300, 20, "There are no ACCT panels known in MAXIS and no additional ACCT information has been added."
				End If
				PushButton 360, 87, 106, 13, "Add New Account Information", add_another_asset_btn
				Text 25, 122, 40, 10, "ACCT - " & acct_count
			End If
			If second_page_display = secu_asset Then								'=====================================================================================   Account - SECU
				GroupBox 70, 90, 405, 190, "Security Account"

				x_pos = 75
				first_acct = TRUE
				for i = 0 to UBound(ASSET_ARRAY)
					If ASSET_ARRAY(i).panel_name = "SECU" Then
						show_acct = FALSE
						If memb_to_match = "" and inst_to_match = "" AND first_acct = TRUE Then
							Text x_pos + 5, 276, 45, 10, "SECU " & ASSET_ARRAY(i).member_ref & " " & ASSET_ARRAY(i).panel_instance
							first_acct = FALSE
							show_acct = TRUE
						ElseIf ASSET_ARRAY(i).member_ref = memb_to_match AND ASSET_ARRAY(i).panel_instance = inst_to_match Then
							Text x_pos + 5, 276, 45, 10, "SECU " & ASSET_ARRAY(i).member_ref & " " & ASSET_ARRAY(i).panel_instance
							show_acct = TRUE
						Else
							PushButton x_pos, 275, 45, 10, "SECU " & ASSET_ARRAY(i).member_ref & " " & ASSET_ARRAY(i).panel_instance, ASSET_ARRAY(i).asset_btn_one
						End If
						' MsgBox "Panel: " & vbCr & "Member - " & ASSET_ARRAY(i).member_ref & vbCr & "Instance - " & ASSET_ARRAY(i).panel_instance & vbCr & "Member to match - " & memb_to_match & vbCr & "Inst to match - " & inst_to_match
						x_pos = x_pos + 45

						If show_acct = TRUE Then
							Text 80, 105, 160, 10, "HH Member: " & ASSET_ARRAY(i).member

							Text 80, 125, 160, 10, "Security Type: " & ASSET_ARRAY(i).asset_type
							Text 80, 140, 160, 10, "Policy/Account Number: " & ASSET_ARRAY(i).account_number
							Text 80, 155, 160, 10, "Security Name: " & ASSET_ARRAY(i).secu_name

							Text 80, 175, 160, 10, "Cash (Surrender) Value: $ " & ASSET_ARRAY(i).secu_cash_value
							Text 80, 190, 160, 10, "As Of: " & ASSET_ARRAY(i).secu_cash_value_date

							Text 250, 190, 160, 10, "Ver: " & ASSET_ARRAY(i).asset_verification
							Text 80, 205, 120, 10, "Face Value of Life Ins: $" & ASSET_ARRAY(i).secu_face_value
							Text 80, 220, 120, 10, "Withdrawal Penalty: $" & ASSET_ARRAY(i).withdraw_penalty
							Text 210, 220, 75, 10, "(Y/N): " & ASSET_ARRAY(i).withdraw_yn
							Text 285, 220, 160, 10, "Ver: " & ASSET_ARRAY(i).withdraw_verif

							Text 80, 240, 160, 10, "Count (Y/N): Cash: " & ASSET_ARRAY(i).count_cash_yn
							Text 240, 240, 50, 10, "FS: " & ASSET_ARRAY(i).count_snap_yn
							Text 290, 240, 50, 10, "HC: " & ASSET_ARRAY(i).count_hc_yn
							Text 340, 240, 50, 10, "GRH: " & ASSET_ARRAY(i).count_grh_yn
							Text 390, 240, 50, 10, "IV-E: " & ASSET_ARRAY(i).count_ive_yn

							Text 80, 255, 160, 10, "Joint Owner (Y/N): " & ASSET_ARRAY(i).joint_owners_yn
							Text 250, 255, 160, 10, "Share Ratio: " & ASSET_ARRAY(i).share_ratio

							PushButton 360, 105, 105, 13, "Update this SECU Panel", update_asset_btn
						End If
					End If
				next
				If secu_count = 0 Then
					Text 110, 105, 300, 20, "There are no SECU panels known in MAXIS and no additional SECU information has been added."
				End If
				PushButton 360, 87, 105, 13, "Add New Security Information", add_another_asset_btn
				Text 25, 137, 40, 10, "SECU - " & secu_count
			End If
			If second_page_display = cars_asset Then								'=====================================================================================   Account - CARS
				GroupBox 70, 90, 405, 190, "Vehicle"

				x_pos = 75
				first_acct = TRUE
				for i = 0 to UBound(ASSET_ARRAY)
					If ASSET_ARRAY(i).panel_name = "CARS" Then
						show_acct = FALSE
						If memb_to_match = "" and inst_to_match = "" AND first_acct = TRUE Then
							Text x_pos + 5, 276, 45, 10, "CARS " & ASSET_ARRAY(i).member_ref & " " & ASSET_ARRAY(i).panel_instance
							first_acct = FALSE
							show_acct = TRUE
						ElseIf ASSET_ARRAY(i).member_ref = memb_to_match AND ASSET_ARRAY(i).panel_instance = inst_to_match Then
							Text x_pos + 5, 276, 45, 10, "CARS " & ASSET_ARRAY(i).member_ref & " " & ASSET_ARRAY(i).panel_instance
							show_acct = TRUE
						Else
							PushButton x_pos, 275, 45, 10, "CARS " & ASSET_ARRAY(i).member_ref & " " & ASSET_ARRAY(i).panel_instance, ASSET_ARRAY(i).asset_btn_one
						End If
						' MsgBox "Panel: " & vbCr & "Member - " & ASSET_ARRAY(i).member_ref & vbCr & "Instance - " & ASSET_ARRAY(i).panel_instance & vbCr & "Member to match - " & memb_to_match & vbCr & "Inst to match - " & inst_to_match
						x_pos = x_pos + 45

						If show_acct = TRUE Then
							Text 80, 105, 160, 10, "HH Member: " & ASSET_ARRAY(i).member

							Text 80, 125, 160, 10, "Vehicle Type: " & ASSET_ARRAY(i).asset_type
							Text 80, 145, 100, 10, "Year: " & ASSET_ARRAY(i).cars_year
							Text 180, 145, 160, 10, "Make: " & ASSET_ARRAY(i).cars_make
							Text 340, 145, 160, 10, "Model: " & ASSET_ARRAY(i).cars_model
							Text 80, 160, 160, 10, "Value     Trade In: $" & ASSET_ARRAY(i).cars_trade_in_value
							Text 200, 160, 160, 10, "Loan: $" & ASSET_ARRAY(i).cars_loan_value
							Text 200, 160, 160, 10, "Source: " & ASSET_ARRAY(i).cars_value_source
							Text 80, 175, 160, 10, "IV-E/Cash/GRH/HC Ownership Ver: " & ASSET_ARRAY(i).asset_verification

							Text 80, 195, 160, 10, "Amount Owed: $" & ASSET_ARRAY(i).cars_amt_owed
							Text 80, 195, 160, 10, "Ver: " & ASSET_ARRAY(i).cars_owed_verification
							Text 80, 210, 160, 10, "As Of: " & ASSET_ARRAY(i).cars_owed_date

							Text 80, 230, 160, 10, "Use: " & ASSET_ARRAY(i).cars_use
							Text 250, 230, 120, 10, "HC Client Benefit (Y/N): " & ASSET_ARRAY(i).cars_hc_benefit
							Text 80, 245, 160, 10, "Joint Owner (Y/N): " & ASSET_ARRAY(i).joint_owners_yn
							Text 250, 245, 160, 10, "Share Ratio: " & ASSET_ARRAY(i).share_ratio

							PushButton 360, 105, 105, 13, "Update This CARS Panel", update_asset_btn
						End If
					End If
				next
				If cars_count = 0 Then
					Text 110, 105, 300, 20, "There are no CARS panels known in MAXIS and no additional CARS information has been added."
				End If
				PushButton 360, 87, 105, 13, "Add New Cars Information", add_another_asset_btn
				Text 25, 152, 40, 10, "CARS - " & cars_count
			End If
			If second_page_display = rest_asset Then								'=====================================================================================   Account - CARS
				GroupBox 70, 90, 405, 190, "Real Estate"

				x_pos = 75
				first_acct = TRUE
				for i = 0 to UBound(ASSET_ARRAY)
					If ASSET_ARRAY(i).panel_name = "REST" Then
						show_acct = FALSE
						If memb_to_match = "" and inst_to_match = "" AND first_acct = TRUE Then
							Text x_pos + 5, 276, 45, 10, "REST " & ASSET_ARRAY(i).member_ref & " " & ASSET_ARRAY(i).panel_instance
							first_acct = FALSE
							show_acct = TRUE
						ElseIf ASSET_ARRAY(i).member_ref = memb_to_match AND ASSET_ARRAY(i).panel_instance = inst_to_match Then
							Text x_pos + 5, 276, 45, 10, "REST " & ASSET_ARRAY(i).member_ref & " " & ASSET_ARRAY(i).panel_instance
							show_acct = TRUE
						Else
							PushButton x_pos, 275, 45, 10, "REST " & ASSET_ARRAY(i).member_ref & " " & ASSET_ARRAY(i).panel_instance, ASSET_ARRAY(i).asset_btn_one
						End If
						' MsgBox "Panel: " & vbCr & "Member - " & ASSET_ARRAY(i).member_ref & vbCr & "Instance - " & ASSET_ARRAY(i).panel_instance & vbCr & "Member to match - " & memb_to_match & vbCr & "Inst to match - " & inst_to_match
						x_pos = x_pos + 45

						If show_acct = TRUE Then
							Text 80, 105, 160, 10, "HH Member: " & ASSET_ARRAY(i).member

							Text 80, 125, 160, 10, "Property Type: " & ASSET_ARRAY(i).asset_type
							Text 250, 125, 160, 10, "Ver: " & ASSET_ARRAY(i).asset_verification
							Text 80, 145, 100, 10, "Market Value: $" & ASSET_ARRAY(i).rest_market_value
							Text 250, 145, 160, 10, "Ver: " & ASSET_ARRAY(i).rest_value_verification
							Text 80, 160, 160, 10, "Amount Owed: $" & ASSET_ARRAY(i).rest_amount_owed
							Text 250, 160, 160, 10, "Ver: " & ASSET_ARRAY(i).rest_owed_verification
							Text 80, 175, 160, 10, "As Of: " & ASSET_ARRAY(i).rest_owed_date

							Text 80, 195, 160, 10, "Property Status: " & ASSET_ARRAY(i).rest_property_status
							Text 80, 210, 160, 10, "Joint Owner (Y/N): " & ASSET_ARRAY(i).joint_owners_yn
							Text 80, 225, 160, 10, "IV-E/Cash/GRH/HC Share Ratio: " & ASSET_ARRAY(i).share_ratio

							Text 80, 245, 160, 10, "IV-E Repayment Agreement Date: " & ASSET_ARRAY(i).rest_ive_repayment_agreement_date

							PushButton 360, 105, 105, 13, "Update This REST Panel", update_asset_btn
						End If
					End If
				next
				If rest_count = 0 Then
					Text 110, 105, 300, 20, "There are no REST panels known in MAXIS and no additional REST information has been added."
				End If
				PushButton 360, 87, 105, 13, "Add New REST Information", add_another_asset_btn
				Text 25, 167, 40, 10, "REST - " & rest_count
			End If

			If second_page_display <> main_asset Then PushButton 20, 90, 40, 15, "MAIN", main_asset_btn
		    If second_page_display <> cash_asset Then PushButton 20, 105, 40, 15, "CASH - " & cash_count, cash_btn
		    If second_page_display <> acct_asset Then PushButton 20, 120, 40, 15, "ACCT - " & acct_count, acct_btn
		    If second_page_display <> secu_asset Then PushButton 20, 135, 40, 15, "SECU - " & secu_count, secu_btn
		    If second_page_display <> cars_asset Then PushButton 20, 150, 40, 15, "CARS - " & cars_count, cars_btn
			If second_page_display <> rest_asset Then PushButton 20, 165, 40, 15, "REST - " & rest_count, rest_btn
		    Text 5, 290, 305, 10, "^^3 - Enter the answers listed on the actual CAF form for Q21 into the 'Answer on the CAF' field."
		    Text 20, 305, 335, 20, "Q. 21. FOR CASH PROGRAMS ONLY: Has anyone in the household given away, sold or traded anything of value in the past 12 months? (For Example: Cash, Bank Accounts, Stocks, Bonds, or Vehicles)?"
		    Text 370, 310, 65, 10, "Answer on the CAF"
		    DropListBox 435, 305, 40, 45, caf_answer_droplist, q21_caf_answer
		    Text 5, 330, 35, 10, "^^4 - ASK - "
		    Text 40, 330, 335, 10, "'Has anyonein the household given away, sold, or traded anythong of value in the past 12 months?'"
			PushButton 365, 325, 115, 15, "YES - Review Improper XFER", improper_transfer_btn
		    Text 40, 350, 70, 10, "Confirm CAF Answer"
		    ComboBox 110, 345, 365, 45, "", q21_confirm_caf_answer


		End If
		If page_display = show_q_22 Then										'THIS IS THE INFORMATION FOR PAGE  ----------------------------------------------------------------------------------------------------------------------------------
			Text 507, 282, 60, 13, "Q. 22"

			Text 5, 10, 305, 10, "^^1 - Enter the answers listed on the actual CAF form for Q5 into the 'Answer on the CAF' field."
			Text 20, 25, 335, 20, "Q. 22. FOR RECERTIFICATIONS ONLY: Did anyone move in or out of your home in the past 12 months?"
			Text 370, 30, 65, 10, "Answer on the CAF"
			DropListBox 435, 25, 40, 45, caf_answer_droplist, q22_caf_answer
			Text 5, 50, 35, 10, "^^2 - ASK - "
			Text 40, 50, 280, 10, "'Did anyone move in or out of your home in the past 12 months?'"
			Text 40, 70, 70, 10, "Confirm CAF Answer"
			ComboBox 110, 65, 365, 45, "", q22_confirm_caf_answer

			Text 5, 90, 300, 10, "^^3 - REVIEW - Review any known members that have left and confirm or add others not listed."

			' Text 20, 110, 150, 10, "Household Members that have left"
			remo_grp_hgt = 15
			for i = 1 to UBound(HH_MEMB_ARRAY, 1)
				If IsDate(HH_MEMB_ARRAY(i).left_hh_date) = TRUE Then
					remo_grp_hgt = remo_grp_hgt + 25
					If HH_MEMB_ARRAY(i).left_hh_expected_return_date <> "" Then remo_grp_hgt = remo_grp_hgt + 10
				End If
			next
			If remo_grp_hgt = 15 Then remo_grp_hgt = 30
			Groupbox 20, 110, 325, remo_grp_hgt, "Household Members that have left"
			y_pos = 125
			member_left_listed = FALSE
			for i = 1 to UBound(HH_MEMB_ARRAY, 1)
				If IsDate(HH_MEMB_ARRAY(i).left_hh_date) = TRUE Then

				 	member_left_listed = TRUE
					Text 25, y_pos, 255, 10, "- MEMB " & HH_MEMB_ARRAY(i).ref_number & "    " & HH_MEMB_ARRAY(i).full_name & " - " & HH_MEMB_ARRAY(i).rel_to_applcnt & " of Memb 01"
					PushButton 290, y_pos, 50, 10, "EDIT", HH_MEMB_ARRAY(i).button_one
					y_pos = y_pos + 10
					Text 50, y_pos, 150, 10, "HH Member Left Date: " & HH_MEMB_ARRAY(i).left_hh_date
					Text 205, y_pos, 150, 10, "Reason: " & HH_MEMB_ARRAY(i).left_hh_reason
					If HH_MEMB_ARRAY(i).left_hh_expected_return_date <> "" Then
						y_pos = y_pos + 10
						Text 50, y_pos, 150, 10, "Expected Return Date: " & HH_MEMB_ARRAY(i).left_hh_expected_return_date
						Text 205, y_pos, 150, 10, "Verif: " & HH_MEMB_ARRAY(i).left_hh_expected_return_verif
					End If
					y_pos = y_pos + 15
				End If
			Next
			If member_left_listed = FALSE Then
				Text 25, y_pos, 255, 10, "No REMO panels in MAXIS or added."
				y_pos = y_pos + 15
			End If
			PushButton 25, y_pos, 150, 10, "Add Another Member that Left", add_remo_btn
			y_pos = y_pos + 20

			Text 5, y_pos, 300, 10, "^^4 - REVIEW - Review any known members that have joined and confirm or add others not listed."
			y_pos = y_pos + 15
			' Text 20, y_pos, 150, 10, "Household Members that have joined:"
			adme_grp_hgt = 15
			for i = 1 to UBound(HH_MEMB_ARRAY, 1)
				If HH_MEMB_ARRAY(i).adme_within_12_months = TRUE Then
					adme_grp_hgt = adme_grp_hgt + 25
				ENd If
			next
			If adme_grp_hgt = 15 Then adme_grp_hgt = 30
			GroupBox 20, y_pos, 325, adme_grp_hgt, "Household Members that have joined"
			y_pos = y_pos + 15

			member_added_listed = FALSE
			for i = 1 to UBound(HH_MEMB_ARRAY, 1)
				If HH_MEMB_ARRAY(i).adme_within_12_months = TRUE Then
					member_added_listed = TRUE

					Text 25, y_pos, 255, 10, "- MEMB " & HH_MEMB_ARRAY(i).ref_number & "    " & HH_MEMB_ARRAY(i).full_name & " - " & HH_MEMB_ARRAY(i).rel_to_applcnt & " of Memb 01"
					PushButton 290, y_pos, 50, 10, "EDIT", HH_MEMB_ARRAY(i).button_one
					y_pos = y_pos + 10
					Text 50, y_pos, 150, 10, "HH Member Arrival Date: " & HH_MEMB_ARRAY(i).adme_arrival_date
					y_pos = y_pos + 15
				End If
			next
			If member_added_listed = FALSE Then
				Text 25, y_pos, 255, 10, "No ADME panels in MAXIS or added."
				y_pos = y_pos + 15
			End If
			PushButton 25, y_pos, 150, 10, "Add Another Member that Joined", add_adme_btn
		End If
		If page_display = show_q_23 Then										'THIS IS THE INFORMATION FOR PAGE  ----------------------------------------------------------------------------------------------------------------------------------
			Text 507, 297, 60, 13, "Q. 23"

			Text 5, 10, 305, 10, "^^1 - Enter the answers listed on the actual CAF form for Q5 into the 'Answer on the CAF' field."
			Text 20, 25, 335, 20, "Q. 23. For children under the age of 19, are both parents living in the home?"
			Text 370, 25, 65, 10, "Answer on the CAF"
			DropListBox 435, 20, 40, 45, caf_answer_droplist, q23_caf_answer
			Text 5, 50, 35, 10, "^^2 - ASK - "
			Text 40, 50, 280, 10, "'For each of the children under the age of 19, are both parents living in the home?"
			Text 40, 70, 70, 10, "Confirm CAF Answer"
			ComboBox 110, 65, 365, 45, "", q23_confirm_caf_answer
			PushButton 2, 100, 18, 10, "^^3", page_23_step_3_btn
			Text 20, 100, 350, 10, " - Review all of the children listed with the client and identify if parents are in or out of the home."
			Text 40, 120, 400, 10, "CHILD'S NAME   -  -  -  -  -  -  -  -  -  - AGE  -  -  PARENTS  -  -  -  -  -  -  -  -  -  -  -  -  -  -  NOTES"

			y_pos = 140
			for i = 0 to UBound(HH_MEMB_ARRAY, 1)
				If HH_MEMB_ARRAY(i).age < 19 Then
					' MsgBox HH_MEMB_ARRAY(i).parents_in_home
					Text 25, y_pos, 100, 10, "M " & HH_MEMB_ARRAY(i).ref_number & " - " & HH_MEMB_ARRAY(i).full_name
					Text 155, y_pos, 30, 20, HH_MEMB_ARRAY(i).age
					DropListBox 185, y_pos-5, 110, 15, "Select One..."+chr(9)+"1 parent in the home"+chr(9)+"Both parents in the home"+chr(9)+"Neither parent in the home"+chr(9)+HH_MEMB_ARRAY(i).parents_in_home, HH_MEMB_ARRAY(i).parents_in_home
					EditBox 305, y_pos-5, 170, 15, HH_MEMB_ARRAY(i).parents_in_home_notes

					Text 50, y_pos+10, 200, 10, "Parent One - " & HH_MEMB_ARRAY(i).parent_one_name
					If HH_MEMB_ARRAY(i).parent_one_in_home = TRUE Then Text 185, y_pos + 10, 100, 10, "Parent is in the home."
					If HH_MEMB_ARRAY(i).parent_one_in_home = FALSE Then Text 185, y_pos + 10, 100, 10, "Absent Parent."

					Text 50, y_pos+20, 200, 10, "Parent Two - " & HH_MEMB_ARRAY(i).parent_two_name
					If HH_MEMB_ARRAY(i).parent_two_in_home = TRUE Then Text 185, y_pos + 20, 100, 10, "Parent is in the home."
					If HH_MEMB_ARRAY(i).parent_two_in_home = FALSE Then Text 185, y_pos + 20, 100, 10, "Absent Parent."

					y_pos = y_pos + 40
				End If
			next




		End If
		If page_display = show_q_24 Then										'THIS IS THE INFORMATION FOR PAGE  ----------------------------------------------------------------------------------------------------------------------------------
			Text 507, 312, 60, 13, "Q. 24"

			Text 5, 10, 305, 10, "^^1 - Enter the answers listed on the actual CAF form for Q24. FOR MSA RECIPIENTS ONLY"
			Text 20, 20, 335, 20, "Q. 24. Does anyone in the household have any of the following expenses?"
			Text 30, 40, 150, 10, "Representative Payee fees"
			DropListBox 135, 35, 40, 45, caf_answer_droplist, rep_payee_fees_caf_answer
			Text 250, 40, 150, 10, "Guardian Conservator fees"
			' Text 350, 40, 170, 10, "(savings, checking, debit card, etc.)"
			DropListBox 355, 35, 40, 45, caf_answer_droplist, guard_consrv_fees_caf_answer

			Text 30, 55, 150, 10, "Physician-prescribed special diet"
			' Text 115, 55, 170, 10, "(stocks, bonds, annuities, 401K, etc.)"
			DropListBox 135, 50, 40, 45, caf_answer_droplist, spec_diet_caf_answer
			Text 250, 55, 150, 10, "High housing costs"
			' Text 350, 52, 130, 20, "(cars, trucks, motorcycles, campers, trailers, etc.)"
			DropListBox 355, 50, 40, 45, caf_answer_droplist, high_housing_caf_answer

			' DIET
			' PDED


			Text 5, 80, 300, 10, "^^2 - CONFIRM DETAILS for each expense type."
			' Text 40, 80, 280, 10, "'Does anyone in the household have any of the following expenses?''"
			' Text 40, 100, 70, 10, "Confirm CAF Answer"
			' ComboBox 110, 95, 365, 45, "", q24_confirm_caf_answer

			y_pos = 115

			Text 30, 100, 200, 10, "Representative Payee fees"
			rep_payee_detail_found = FALSE
			for i = 0 to UBound(HH_MEMB_ARRAY, 1)
				If HH_MEMB_ARRAY(i).pded_rep_payee_fee <> "" Then
					rep_payee_detail_found = TRUE

					Text 50, y_pos, 200, 10, "Rep Payee Fees for M " & HH_MEMB_ARRAY(i).ref_number & " - " & HH_MEMB_ARRAY(i).full_name
					Text 250, y_pos, 5, 10, "$"
					EditBox 260, y_pos-5, 75, 15, HH_MEMB_ARRAY(i).pded_rep_payee_fee
					y_pos = y_pos + 15
				End If
			Next
			If rep_payee_detail_found = FALSE Then
				Text 50, y_pos, 300, 10, "No Rep Payee Fees listed for any member."
				y_pos = y_pos + 15
			End If
			y_pos = y_pos + 5

			Text 30, y_pos, 200, 10, "Guardian or Conservator fees"
			y_pos = y_pos + 15
			guardian_fee_found = FALSE
			for i = 0 to UBound(HH_MEMB_ARRAY, 1)
				If HH_MEMB_ARRAY(i).pded_guardian_fee <> "" Then
					guardian_fee_found = TRUE

					Text 50, y_pos, 200, 10, "Guardian Fees for M " & HH_MEMB_ARRAY(i).ref_number & " - " & HH_MEMB_ARRAY(i).full_name
					Text 250, y_pos, 5, 10, "$"
					EditBox 260, y_pos-5, 75, 15, HH_MEMB_ARRAY(i).pded_guardian_fee
					y_pos = y_pos + 15
				End If
			Next
			If guardian_fee_found = FALSE Then
				Text 50, y_pos, 300, 10, "No Guardian or Conservator Fees listed for any member."
				y_pos = y_pos + 15
			End If
			y_pos = y_pos + 5

			Text 30, y_pos, 200, 10, "Physician-prescribed special diet"
			y_pos = y_pos + 15
			special_diet_found = FALSE
			for i = 0 to UBound(HH_MEMB_ARRAY, 1)
				If HH_MEMB_ARRAY(i).diet_exists = TRUE Then
					special_diet_found = TRUE

					Text 50, y_pos, 250, 10, "M " & HH_MEMB_ARRAY(i).ref_number & " - " & HH_MEMB_ARRAY(i).full_name & " special diet(s):"
					y_pos = y_pos + 15

					If HH_MEMB_ARRAY(i).diet_msa_type_one <> "" Then
						Text 60, y_pos, 150, 10, "Diet Type: " & HH_MEMB_ARRAY(i).diet_msa_type_one
						Text 215, y_pos, 150, 10, "Verified: " & HH_MEMB_ARRAY(i).diet_msa_verif_one
						y_pos = y_pos + 15
					End If
					If HH_MEMB_ARRAY(i).diet_msa_type_two <> "" Then
						Text 60, y_pos, 150, 10, "Diet Type: " & HH_MEMB_ARRAY(i).diet_msa_type_two
						Text 215, y_pos, 150, 10, "Verified: " & HH_MEMB_ARRAY(i).diet_msa_verif_two
						y_pos = y_pos + 15
					End If
					If HH_MEMB_ARRAY(i).diet_msa_type_three <> "" Then
						Text 60, y_pos, 150, 10, "Diet Type: " & HH_MEMB_ARRAY(i).diet_msa_type_three
						Text 215, y_pos, 150, 10, "Verified: " & HH_MEMB_ARRAY(i).diet_msa_verif_three
						y_pos = y_pos + 15
					End If
					If HH_MEMB_ARRAY(i).diet_msa_type_four <> "" Then
						Text 60, y_pos, 150, 10, "Diet Type: " & HH_MEMB_ARRAY(i).diet_msa_type_four
						Text 215, y_pos, 150, 10, "Verified: " & HH_MEMB_ARRAY(i).diet_msa_verif_four
						y_pos = y_pos + 15
					End If
					If HH_MEMB_ARRAY(i).diet_msa_type_five <> "" Then
						Text 60, y_pos, 150, 10, "Diet Type: " & HH_MEMB_ARRAY(i).diet_msa_type_five
						Text 215, y_pos, 150, 10, "Verified: " & HH_MEMB_ARRAY(i).diet_msa_verif_five
						y_pos = y_pos + 15
					End If
					If HH_MEMB_ARRAY(i).diet_msa_type_six <> "" Then
						Text 60, y_pos, 150, 10, "Diet Type: " & HH_MEMB_ARRAY(i).diet_msa_type_six
						Text 215, y_pos, 150, 10, "Verified: " & HH_MEMB_ARRAY(i).diet_msa_verif_six
						y_pos = y_pos + 15
					End If
					If HH_MEMB_ARRAY(i).diet_msa_type_seven <> "" Then
						Text 60, y_pos, 150, 10, "Diet Type: " & HH_MEMB_ARRAY(i).diet_msa_type_seven
						Text 215, y_pos, 150, 10, "Verified: " & HH_MEMB_ARRAY(i).diet_msa_verif_seven
						y_pos = y_pos + 15
					End If
					If HH_MEMB_ARRAY(i).diet_msa_type_eight <> "" Then
						Text 60, y_pos, 150, 10, "Diet Type: " & HH_MEMB_ARRAY(i).diet_msa_type_eight
						Text 215, y_pos, 150, 10, "Verified: " & HH_MEMB_ARRAY(i).diet_msa_verif_eight
						y_pos = y_pos + 15
					End If
				End If
			next
			If special_diet_found = FALSE Then
				Text 50, y_pos, 300, 10, "No Special Diet listed for any member."
				y_pos = y_pos + 15
			End If
			y_pos = y_pos + 5

			Text 30, y_pos, 200, 10, "High housing costs"
			y_pos = y_pos + 15
			shel_needy_found = FALSE
			for i = 0 to UBound(HH_MEMB_ARRAY, 1)
				If HH_MEMB_ARRAY(i).pded_shel_spec_need <> "" Then
					shel_needy_found = TRUE

					If HH_MEMB_ARRAY(i).pded_shel_spec_need = "Yes" Then Text 50, y_pos, 300, 10, "M " & HH_MEMB_ARRAY(i).ref_number & " - " & HH_MEMB_ARRAY(i).full_name & " indicated as 'SHELTER NEEDY'"
					If HH_MEMB_ARRAY(i).pded_shel_spec_need = "No" Then Text 50, y_pos, 300, 10, "M " & HH_MEMB_ARRAY(i).ref_number & " - " & HH_MEMB_ARRAY(i).full_name & " NO shelter special need."
					y_pos = y_pos + 15
				End If
			next
			If shel_needy_found = FALSE Then
				Text 50, y_pos, 300, 10, "No Shelter Special Need listed for any member."
				y_pos = y_pos + 15
			End If
			y_pos = y_pos + 10

			PushButton 355, y_pos, 100, 15, "EDIT EXPENSE DETAIL", edit_expense_detail_btn

			' Text 495, 267, 60, 15, "Q. 20 and 21"
			' Text 5, 10, 200, 10, "^^1 - Enter the answers listed on the actual CAF form for Q20"
			' Text 20, 20, 335, 10, "Q. 20. Does anyone in the household own, or is anyone buying, any of the following?"
			' Text 30, 40, 20, 10, "Cash"
			' DropListBox 70, 35, 40, 45, caf_answer_droplist, cash_caf_answer
			' Text 250, 40, 55, 10, "Bank accounts"
			' Text 350, 40, 170, 10, "(savings, checking, debit card, etc.)"
			' DropListBox 305, 35, 40, 45, caf_answer_droplist, acct_caf_answer
			'
			' Text 30, 55, 40, 10, "Securities"
			' Text 115, 55, 170, 10, "(stocks, bonds, annuities, 401K, etc.)"
			' DropListBox 70, 50, 40, 45, caf_answer_droplist, secu_caf_answer
			' Text 250, 55, 35, 10, "Vehicles"
			' Text 350, 52, 130, 20, "(cars, trucks, motorcycles, campers, trailers, etc.)"
			' DropListBox 305, 50, 40, 45, caf_answer_droplist, cars_caf_answer

		End If
		If page_display = show_qual Then										'THIS IS THE INFORMATION FOR PAGE  ----------------------------------------------------------------------------------------------------------------------------------
			Text 492, 327, 60, 13, "CAF QUAL Q"

			Text 10, 10, 350, 10, "^^1 - ENTER the answers listed on the actual CAF on page 8."

			Text 10, 30, 395, 15, "Qualifying Questions are listed at the end of the CAF form and are completed by the client. Indicate the answers to those questions here. If any are 'Yes' then indicate which household member to which the question refers."
			Text 15, 60, 395, 20, "Has a court or any other civil or administrative process in Minnesota or any other state found anyone in the household guilty or has anyone been disqualified from receiving public assistance for breaking any of the rules listed in the CAF?"
			DropListBox 415, 60, 40, 45, caf_answer_droplist, qual_question_one
			Text 15, 85, 395, 20, "Has anyone in the household been convicted of making fraudulent statements about their place of residence to get cash or SNAP benefits from more than one state?"
			DropListBox 415, 85, 40, 45, caf_answer_droplist, qual_question_two
			Text 15, 110, 395, 20, "Is anyone in your householdhiding or running from the law to avoid prosecution being taken into custody, or to avoid going to jail for a felony?"
			DropListBox 415, 110, 40, 45, caf_answer_droplist, qual_question_three
			Text 15, 135, 395, 20, "Has anyone in your household been convicted of a drug felony in the past 10 years?"
			DropListBox 415, 135, 40, 45, caf_answer_droplist, qual_question_four
			Text 15, 150, 395, 20, "Is anyone in your household currently violating a condition of parole, probation or supervised release?"
			DropListBox 415, 150, 40, 45, caf_answer_droplist, qual_question_five

			Text 10, 170, 350, 10, "^^2 - CONFIRM verbally with the client the answers provided above."
			Text 40, 190, 70, 10, "Confirm CAF Answer"
			ComboBox 110, 185, 365, 45, "", qual_q_confirm_caf_answer

			Text 10, 210, 350, 10, "^^3 - If ANY questions have been answerd as 'YES', enter the quesiton number and HH Member to which the question applies."

			y_pos = 230
			For qual_yes = 0 to UBound(QUAL_QUESTIONS_ARRAY,2)
				Text 40, y_pos, 45, 10, "Question No."
				DropListBox 90, y_pos-5, 55, 15, "Select"+chr(9)+"Q. 1"+chr(9)+"Q. 2"+chr(9)+"Q. 3"+chr(9)+"Q. 4"+chr(9)+"Q. 5", QUAL_QUESTIONS_ARRAY(quest_numb, qual_yes)
				Text 160, y_pos, 40, 10, "HH Member"
				ComboBox 205, y_pos-5, 105, 45, memb_droplist, QUAL_QUESTIONS_ARRAY(memb_numb, qual_yes)
				y_pos = y_pos + 15
			Next
			PushButton 40, y_pos, 110, 10, "ADD ANOTHER 'YES' DETAIL", add_qual_quest_yes
			' Text 260, 40, 70, 10, "Household Member:"
			' Text 260, 80, 70, 10, "Household Member:"
			' Text 260, 110, 70, 10, "Household Member:"
			' Text 260, 140, 70, 10, "Household Member:"
			' Text 260, 160, 70, 10, "Household Member:"
			' ComboBox 340, 40, 105, 45, memb_droplist, qual_memb_one
			' ComboBox 340, 80, 105, 45, memb_droplist, qual_memb_two
			' ComboBox 340, 110, 105, 45, memb_droplist, qual_memb_there
			' ComboBox 340, 140, 105, 45, memb_droplist, qual_memb_four
			' ComboBox 340, 160, 105, 45, memb_droplist, qual_memb_five
		End If
		If page_display = show_pg_last Then										'THIS IS THE INFORMATION FOR PAGE  ----------------------------------------------------------------------------------------------------------------------------------
			Text 490, 342, 60, 13, "CAF Last Page"
		End If

		If page_display <> show_pg_one Then PushButton 485, 10, 60, 13, "CAF Page 1", caf_page_one_btn
		If page_display <> show_pg_memb_list AND page_display <> show_pg_memb_info Then PushButton 485, 25, 60, 13, "CAF MEMBs", caf_membs_btn
		If page_display <> show_q_1_2 Then PushButton 485, 40, 60, 13, "Q. 1 and 2", caf_q_1_2_btn
		If page_display <> show_q_3 Then PushButton 485, 55, 60, 13, "Q. 3", caf_q_3_btn
		If page_display <> show_q_4 Then PushButton 485, 70, 60, 13, "Q. 4", caf_q_4_btn
		If page_display <> show_q_5 Then PushButton 485, 85, 60, 13, "Q. 5", caf_q_5_btn
		If page_display <> show_q_6 Then PushButton 485, 100, 60, 13, "Q. 6", caf_q_6_btn
		If page_display <> show_q_7 Then PushButton 485, 115, 60, 13, "Q. 7", caf_q_7_btn
		If page_display <> show_q_8 Then PushButton 485, 130, 60, 13, "Q. 8", caf_q_8_btn
		If page_display <> show_q_9 Then PushButton 485, 145, 60, 13, "Q. 9", caf_q_9_btn
		If page_display <> show_q_10 Then PushButton 485, 160, 60, 13, "Q. 10", caf_q_10_btn
		If page_display <> show_q_11 Then PushButton 485, 175, 60, 13, "Q. 11", caf_q_11_btn
		If page_display <> show_q_12 Then PushButton 485, 190, 60, 13, "Q. 12", caf_q_12_btn
		If page_display <> show_q_13 Then PushButton 485, 205, 60, 13, "Q. 13", caf_q_13_btn
		If page_display <> show_q_14_15 Then PushButton 485, 220, 60, 13, "Q. 14 and 15", caf_q_14_15_btn
		If page_display <> show_q_16_18 Then PushButton 485, 235, 60, 13, "Q. 16, 17, and 18", caf_q_16_17_18_btn
		If page_display <> show_q_19 Then PushButton 485, 250, 60, 13, "Q. 19", caf_q_19_btn
		If page_display <> show_q_20_21 Then PushButton 485, 265, 60, 13, "Q. 20 and 21", caf_q_20_21_btn
		If page_display <> show_q_22 Then PushButton 485, 280, 60, 13, "Q. 22", caf_q_22_btn
		If page_display <> show_q_23 Then PushButton 485, 295, 60, 13, "Q. 23", caf_q_23_btn
		If page_display <> show_q_24 Then PushButton 485, 310, 60, 13, "Q. 24", caf_q_24_btn
		If page_display <> show_qual Then PushButton 485, 325, 60, 13, "CAF QUAL Q", caf_qual_q_btn
		If page_display <> show_pg_last Then PushButton 485, 340, 60, 13, "CAF Last Page", caf_last_page_btn
		PushButton 415, 365, 50, 15, "NEXT", next_btn
		PushButton 465, 365, 80, 15, "Complete Interview", finish_interview_btn



	EndDialog
end function

'CAF PAGE 1 - HH Comp and Address
' FUNCTION - Read the current address on ADDR - resi and mail and phone numbers.
function read_ADDR_panel(addr_eff_date, line_one, line_two, city, state, zip, county, verif, homeless, ind_reservation, living_sit, res_name, mail_line_one, mail_line_two, mail_city, mail_state, mail_zip, phone_one, type_one, phone_two, type_two, phone_three, type_three, updated_date)
    Call navigate_to_MAXIS_screen("STAT", "ADDR")

    EMReadScreen line_one, 22, 6, 43
    EMReadScreen line_two, 22, 7, 43
    EMReadScreen city_line, 15, 8, 43
    EMReadScreen state_line, 2, 8, 66
    EMReadScreen zip_line, 7, 9, 43
    EMReadScreen county_line, 2, 9, 66
    EMReadScreen verif_line, 2, 9, 74
    EMReadScreen homeless, 1, 10, 43
    EMReadScreen ind_reservation, 1, 10, 74
    EMReadScreen living_sit, 2, 11, 43
	EMReadScreen res_name, 2, 11, 74

    line_one = replace(line_one, "_", "")
    line_two = replace(line_two, "_", "")
    city = replace(city_line, "_", "")
    state = state_line
    zip = replace(zip_line, "_", "")

    If county_line = "01" Then county = "01 - Aitkin"
    If county_line = "02" Then county = "02 - Anoka"
    If county_line = "03" Then county = "03 - Becker"
    If county_line = "04" Then county = "04 - Beltrami"
    If county_line = "05" Then county = "05 - Benton"
    If county_line = "06" Then county = "06 - Big Stone"
    If county_line = "07" Then county = "07 - Blue Earth"
    If county_line = "08" Then county = "08 - Brown"
    If county_line = "09" Then county = "09 - Carlton"
    If county_line = "10" Then county = "10 - Carver"
    If county_line = "11" Then county = "11 - Cass"
    If county_line = "12" Then county = "12 - Chippewa"
    If county_line = "13" Then county = "13 - Chisago"
    If county_line = "14" Then county = "14 - Clay"
    If county_line = "15" Then county = "15 - Clearwater"
    If county_line = "16" Then county = "16 - Cook"
    If county_line = "17" Then county = "17 - Cottonwood"
    If county_line = "18" Then county = "18 - Crow Wing"
    If county_line = "19" Then county = "19 - Dakota"
    If county_line = "20" Then county = "20 - Dodge"
    If county_line = "21" Then county = "21 - Douglas"
    If county_line = "22" Then county = "22 - Faribault"
    If county_line = "23" Then county = "23 - Fillmore"
    If county_line = "24" Then county = "24 - Freeborn"
    If county_line = "25" Then county = "25 - Goodhue"
    If county_line = "26" Then county = "26 - Grant"
    If county_line = "27" Then county = "27 - Hennepin"
    If county_line = "28" Then county = "28 - Houston"
    If county_line = "29" Then county = "29 - Hubbard"
    If county_line = "30" Then county = "30 - Isanti"
    If county_line = "31" Then county = "31 - Itasca"
    If county_line = "32" Then county = "32 - Jackson"
    If county_line = "33" Then county = "33 - Kanabec"
    If county_line = "34" Then county = "34 - Kandiyohi"
    If county_line = "35" Then county = "35 - Kittson"
    If county_line = "36" Then county = "36 - Koochiching"
    If county_line = "37" Then county = "37 - Lac Qui Parle"
    If county_line = "38" Then county = "38 - Lake"
    If county_line = "39" Then county = "39 - Lake Of Woods"
    If county_line = "40" Then county = "40 - Le Sueur"
    If county_line = "41" Then county = "41 - Lincoln"
    If county_line = "42" Then county = "42 - Lyon"
    If county_line = "43" Then county = "43 - Mcleod"
    If county_line = "44" Then county = "44 - Mahnomen"
    If county_line = "45" Then county = "45 - Marshall"
    If county_line = "46" Then county = "46 - Martin"
    If county_line = "47" Then county = "47 - Meeker"
    If county_line = "48" Then county = "48 - Mille Lacs"
    If county_line = "49" Then county = "49 - Morrison"
    If county_line = "50" Then county = "50 - Mower"
    If county_line = "51" Then county = "51 - Murray"
    If county_line = "52" Then county = "52 - Nicollet"
    If county_line = "53" Then county = "53 - Nobles"
    If county_line = "54" Then county = "54 - Norman"
    If county_line = "55" Then county = "55 - Olmsted"
    If county_line = "56" Then county = "56 - Otter Tail"
    If county_line = "57" Then county = "57 - Pennington"
    If county_line = "58" Then county = "58 - Pine"
    If county_line = "59" Then county = "59 - Pipestone"
    If county_line = "60" Then county = "60 - Polk"
    If county_line = "61" Then county = "61 - Pope"
    If county_line = "62" Then county = "62 - Ramsey"
    If county_line = "63" Then county = "63 - Red Lake"
    If county_line = "64" Then county = "64 - Redwood"
    If county_line = "65" Then county = "65 - Renville"
    If county_line = "66" Then county = "66 - Rice"
    If county_line = "67" Then county = "67 - Rock"
    If county_line = "68" Then county = "68 - Roseau"
    If county_line = "69" Then county = "69 - St. Louis"
    If county_line = "70" Then county = "70 - Scott"
    If county_line = "71" Then county = "71 - Sherburne"
    If county_line = "72" Then county = "72 - Sibley"
    If county_line = "73" Then county = "73 - Stearns"
    If county_line = "74" Then county = "74 - Steele"
    If county_line = "75" Then county = "75 - Stevens"
    If county_line = "76" Then county = "76 - Swift"
    If county_line = "77" Then county = "77 - Todd"
    If county_line = "78" Then county = "78 - Traverse"
    If county_line = "79" Then county = "79 - Wabasha"
    If county_line = "80" Then county = "80 - Wadena"
    If county_line = "81" Then county = "81 - Waseca"
    If county_line = "82" Then county = "82 - Washington"
    If county_line = "83" Then county = "83 - Watonwan"
    If county_line = "84" Then county = "84 - Wilkin"
    If county_line = "85" Then county = "85 - Winona"
    If county_line = "86" Then county = "86 - Wright"
    If county_line = "87" Then county = "87 - Yellow Medicine"
    If county_line = "89" Then county = "89 - Out-of-State"

    If homeless = "Y" Then homeless = "Yes"
    If homeless = "N" Then homeless = "No"
    If ind_reservation = "Y" Then ind_reservation = "Yes"
    If ind_reservation = "N" Then ind_reservation = "No"

    If verif_line = "SF" Then verif = "SF - Shelter Form"
    If verif_line = "Co" Then verif = "CO - Coltrl Stmt"
    If verif_line = "MO" Then verif = "MO - Mortgage Papers"
    If verif_line = "TX" Then verif = "TX - Prop Tax Stmt"
    If verif_line = "CD" Then verif = "CD - Contrct for Deed"
    If verif_line = "UT" Then verif = "UT - Utility Stmt"
    If verif_line = "DL" Then verif = "DL - Driver Lic/State ID"
    If verif_line = "OT" Then verif = "OT - Other Document"
    If verif_line = "NO" Then verif = "NO - No Ver Prvd"
    If verif_line = "?_" Then verif = "? - Delayed"
    If verif_line = "__" Then verif = "Blank"


    If living_sit = "__" Then living_sit = "Blank"
    If living_sit = "01" Then living_sit = "01 - Own Housing (lease, mortgage, or roomate)"
    If living_sit = "02" Then living_sit = "02 - Family/Friends due to economic hardship"
    If living_sit = "03" Then living_sit = "03 - Servc prvdr- foster/group home"
    If living_sit = "04" Then living_sit = "04 - Hospital/Treatment/Detox/Nursing Home"
    If living_sit = "05" Then living_sit = "05 - Jail/Prison//Juvenile Det."
    If living_sit = "06" Then living_sit = "06 - Hotel/Motel"
    If living_sit = "07" Then living_sit = "07 - Emergency Shelter"
    If living_sit = "08" Then living_sit = "08 - Place not meant for Housing"
    If living_sit = "09" Then living_sit = "09 - Declined"
    If living_sit = "10" Then living_sit = "10 - Unknown"

	If res_name = "__" Then res_name = "Blank"
	If res_name = "BD" Then res_name = "Bois Forte - Deer Creek"
	If res_name = "BN" Then res_name = "Bois Forte - Nett Lake"
	If res_name = "BV" Then res_name = "Bois Forte - Vermillion Lk"
	If res_name = "FL" Then res_name = "Fond du Lac"
	If res_name = "GP" Then res_name = "Grand Portage"
	If res_name = "LL" Then res_name = "Leach Lake"
	If res_name = "LS" Then res_name = "Lower Sioux"
	If res_name = "ML" Then res_name = "Mille Lacs"
	If res_name = "PL" Then res_name = "Prairie Island Community"
	If res_name = "RL" Then res_name = "Red Lake"
	If res_name = "SM" Then res_name = "Shakopee Mdewakanton"
	If res_name = "US" Then res_name = "Upper Sioux"
	If res_name = "WE" Then res_name = "White Earth"

    EMReadScreen addr_eff_date, 8, 4, 43
    EMReadScreen addr_future_date, 8, 4, 66
    EMReadScreen mail_line_one, 22, 13, 43
    EMReadScreen mail_line_two, 22, 14, 43
    EMReadScreen mail_city, 15, 15, 43
    EMReadScreen mail_state, 2, 16, 43
    EMReadScreen mail_zip, 7, 16, 52

    addr_eff_date = replace(addr_eff_date, " ", "/")
    addr_future_date = trim(addr_future_date)
    addr_future_date = replace(addr_future_date, " ", "/")
    mail_line_one = replace(mail_line_one, "_", "")
    mail_line_two = replace(mail_line_two, "_", "")
    mail_city = replace(mail_city, "_", "")
    mail_state = replace(mail_state, "_", "")
    mail_zip = replace(mail_zip, "_", "")

	EMReadScreen phone_one, 14, 17, 45
	EMReadScreen phone_two, 14, 18, 45
	EMReadScreen phone_three, 14, 19, 45
	EMReadScreen type_one, 1, 17, 67
	EMReadScreen type_two, 1, 18, 67
	EMReadScreen type_three, 1, 19, 67

	phone_one = "(" & replace(replace(replace(phone_one, " ) ", ")"), " ", " - "), ")", ") ")
	If phone_one = "(___) ___ - ____" Then phone_one = ""
	If type_one = "_" Then type_one = "Unknown"
	If type_one = "H" Then type_one = "Home"
	If type_one = "W" Then type_one = "Work"
	If type_one = "C" Then type_one = "Cell"
	If type_one = "M" Then type_one = "Message"
	If type_one = "T" Then type_one = "TTY/TDD"

	phone_two = "(" & replace(replace(replace(phone_two, " ) ", ")"), " ", " - "), ")", ") ")
	If phone_two = "(___) ___ - ____" Then phone_two = ""
	If type_two = "_" Then type_two = "Unknown"
	If type_two = "H" Then type_two = "Home"
	If type_two = "W" Then type_two = "Work"
	If type_two = "C" Then type_two = "Cell"
	If type_two = "M" Then type_two = "Message"
	If type_two = "T" Then type_two = "TTY/TDD"

	phone_three = "(" & replace(replace(replace(phone_three, " ) ", ")"), " ", " - "), ")", ") ")
	If phone_three = "(___) ___ - ____" Then phone_three = ""
	If type_three = "_" Then type_three = "Unknown"
	If type_three = "H" Then type_three = "Home"
	If type_three = "W" Then type_three = "Work"
	If type_three = "C" Then type_three = "Cell"
	If type_three = "M" Then type_three = "Message"
	If type_three = "T" Then type_three = "TTY/TDD"

	EMReadScreen updated_date, 8, 21, 55
	updated_date = replace(updated_date, " ", "/")
end function


' function read_all_COEX()
' 	Call navigate_to_MAXIS_screen ("STAT", "PNLE")
'
' 	pnle_row = 3
' 	count = 1
' 	previous_member = ""
' 	panel_count = UBound(INCOME_ARRAY, 1)
' 	start_count = panel_count
' 	coex_found = FALSE
' 	Do
' 		EMReadScreen panel_name, 4, pnle_row, 5
' 		' MsgBox panel_name
' 		IF panel_name = "UNEA" Then
' 			EMReadScreen panel_memb, 2, pnle_row, 10
' 			If panel_memb <> previous_member Then count = 1
'
' 			ReDim Preserve INCOME_ARRAY(panel_count)
' 			Set INCOME_ARRAY(panel_count) = new client_income
' 			INCOME_ARRAY(panel_count).member_ref = panel_memb
' 			INCOME_ARRAY(panel_count).panel_instance = "0" & count
'
'
' 			panel_count = panel_count + 1
' 			count = count + 1
' 			previous_member = panel_memb
' 			' MsgBox panel_count
' 			coex_found = TRUE
' 			case_has_income_listed = TRUE
'
' 		End If
' 		pnle_row = pnle_row + 1
' 		If pnle_row = 20 Then
' 			transmit
' 			pnle_row = 3
' 		End If
' 		EMReadScreen panel_summ, 4, 2, 53
' 		' MsgBox "PNLI Row - " & pnle_row & vbNewLine & "SUMM - " & panel_summ
' 	Loop until panel_summ = "PNLE"
' 	' If count = 1 Then ReDim Preserve INCOME_ARRAY(panel_count)
' 	stop_count = panel_count - 1
' 	If stop_count < 0 Then stop_count = 0
' end function

function read_all_DCEX()
end function

' FUNCTION - Read all the HH Members from case, including how they are related to M01 and SIBL and PARE - get age
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

	For the_rela = 0 to UBound(ALL_HH_RELATIONSHIPS_ARRAY, 2)
		ALL_HH_RELATIONSHIPS_ARRAY(rela_pers_one, the_rela) = ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_ref, the_rela) & " - " & ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_one_name, the_rela)
		ALL_HH_RELATIONSHIPS_ARRAY(rela_pers_two, the_rela) = ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_ref, the_rela) & " - " & ALL_HH_RELATIONSHIPS_ARRAY(rela_clt_two_name, the_rela)

		' MsgBox "Relationship detail:" & vbNewLine & ALL_HH_RELATIONSHIPS_ARRAY(rela_pers_one, the_rela) & " is the " & ALL_HH_RELATIONSHIPS_ARRAY(rela_type, the_rela) & " of " & ALL_HH_RELATIONSHIPS_ARRAY(rela_pers_two, the_rela)
	Next
end function

Dim ASSET_ARRAY()
ReDim ASSET_ARRAY(0)
function read_all_Assets()
	Call navigate_to_MAXIS_screen ("STAT", "PNLR")

	pnlr_row = 3
	acct_inst = 1
	secu_inst = 1
	cars_inst = 1
	rest_inst = 1
	previous_member = ""
	start_count = 0
	asset_count = 0
	Do
		EMReadScreen panel_name, 4, pnlr_row, 5
		' MsgBox panel_name
		IF panel_name = "CASH" Then
			EMReadScreen panel_memb, 2, pnlr_row, 10

			ReDim Preserve ASSET_ARRAY(asset_count)
			Set ASSET_ARRAY(asset_count) = new client_assets
			ASSET_ARRAY(asset_count).member_ref = panel_memb
			ASSET_ARRAY(asset_count).panel_instance = "01"
			ASSET_ARRAY(asset_count).panel_name = "CASH"
			ASSET_ARRAY(asset_count).asset_btn_one = 700 + asset_count

			asset_count = asset_count + 1
			cash_count = cash_count + 1

			' MsgBox asset_count
			unea_found = TRUE
		End If

		IF panel_name = "ACCT" Then
			EMReadScreen panel_memb, 2, pnlr_row, 10
			If panel_memb <> previous_member Then acct_inst = 1

			ReDim Preserve ASSET_ARRAY(asset_count)
			Set ASSET_ARRAY(asset_count) = new client_assets
			ASSET_ARRAY(asset_count).member_ref = panel_memb
			ASSET_ARRAY(asset_count).panel_instance = "0" & acct_inst
			ASSET_ARRAY(asset_count).panel_name = "ACCT"
			ASSET_ARRAY(asset_count).asset_btn_one = 700 + asset_count

			asset_count = asset_count + 1
			acct_inst = acct_inst + 1
			acct_count = acct_count + 1
			previous_member = panel_memb
			' MsgBox asset_count
			unea_found = TRUE
		End If

		IF panel_name = "SECU" Then
			EMReadScreen panel_memb, 2, pnlr_row, 10
			If panel_memb <> previous_member Then secu_inst = 1

			ReDim Preserve ASSET_ARRAY(asset_count)
			Set ASSET_ARRAY(asset_count) = new client_assets
			ASSET_ARRAY(asset_count).member_ref = panel_memb
			ASSET_ARRAY(asset_count).panel_instance = "0" & secu_inst
			ASSET_ARRAY(asset_count).panel_name = "SECU"
			ASSET_ARRAY(asset_count).asset_btn_one = 700 + asset_count

			asset_count = asset_count + 1
			secu_inst = secu_inst + 1
			secu_count = secu_count + 1
			previous_member = panel_memb
			' MsgBox asset_count
			unea_found = TRUE
		End If

		IF panel_name = "CARS" Then
			EMReadScreen panel_memb, 2, pnlr_row, 10
			If panel_memb <> previous_member Then cars_inst = 1

			ReDim Preserve ASSET_ARRAY(asset_count)
			Set ASSET_ARRAY(asset_count) = new client_assets
			ASSET_ARRAY(asset_count).member_ref = panel_memb
			ASSET_ARRAY(asset_count).panel_instance = "0" & cars_inst
			ASSET_ARRAY(asset_count).panel_name = "CARS"
			ASSET_ARRAY(asset_count).asset_btn_one = 700 + asset_count

			asset_count = asset_count + 1
			cars_inst = cars_inst + 1
			cars_count = cars_count + 1
			previous_member = panel_memb
			' MsgBox asset_count
			unea_found = TRUE
		End If

		IF panel_name = "REST" Then
			EMReadScreen panel_memb, 2, pnlr_row, 10
			If panel_memb <> previous_member Then rest_inst = 1

			ReDim Preserve ASSET_ARRAY(asset_count)
			Set ASSET_ARRAY(asset_count) = new client_assets
			ASSET_ARRAY(asset_count).member_ref = panel_memb
			ASSET_ARRAY(asset_count).panel_instance = "0" & rest_inst
			ASSET_ARRAY(asset_count).panel_name = "REST"
			ASSET_ARRAY(asset_count).asset_btn_one = 700 + asset_count

			asset_count = asset_count + 1
			rest_inst = rest_inst + 1
			rest_count = rest_count + 1
			previous_member = panel_memb
			' MsgBox asset_count
			unea_found = TRUE
		End If
		pnlr_row = pnlr_row + 1
		If pnlr_row = 20 Then
			transmit
			pnlr_row = 3
		End If
		EMReadScreen panel_summ, 4, 2, 53
	Loop until panel_summ = "PNLI"

	If asset_count = 0 Then
		Set ASSET_ARRAY(0) = new client_assets
	End If


	If unea_found = TRUE Then
		for i = 0 to UBound(ASSET_ARRAY)
			ASSET_ARRAY(i).read_member_name

			If ASSET_ARRAY(i).panel_name = "CASH" Then ASSET_ARRAY(i).read_cash_panel
			If ASSET_ARRAY(i).panel_name = "ACCT" Then ASSET_ARRAY(i).read_acct_panel
			If ASSET_ARRAY(i).panel_name = "SECU" Then ASSET_ARRAY(i).read_secu_panel
			If ASSET_ARRAY(i).panel_name = "CARS" Then ASSET_ARRAY(i).read_cars_panel
			If ASSET_ARRAY(i).panel_name = "REST" Then ASSET_ARRAY(i).read_rest_panel
			' MsgBox ASSET_ARRAY(i).panel_name
		Next

	End If
end function

function read_all_UNEA()
	Call navigate_to_MAXIS_screen ("STAT", "PNLI")

	pnli_row = 3
	count = 1
	previous_member = ""
	panel_count = UBound(INCOME_ARRAY, 1)
	start_count = panel_count
	unea_found = FALSE
	Do
		EMReadScreen panel_name, 4, pnli_row, 5
		' MsgBox panel_name
		IF panel_name = "UNEA" Then
			EMReadScreen panel_memb, 2, pnli_row, 10
			If panel_memb <> previous_member Then count = 1

			ReDim Preserve INCOME_ARRAY(panel_count)
			Set INCOME_ARRAY(panel_count) = new client_income
			INCOME_ARRAY(panel_count).member_ref = panel_memb
			INCOME_ARRAY(panel_count).panel_instance = "0" & count

			panel_count = panel_count + 1
			count = count + 1
			previous_member = panel_memb
			' MsgBox panel_count
			unea_found = TRUE
			case_has_income_listed = TRUE

		End If
		pnli_row = pnli_row + 1
		If pnli_row = 20 Then
			transmit
			pnli_row = 3
		End If
		EMReadScreen panel_summ, 4, 2, 53
		' MsgBox "PNLI Row - " & pnli_row & vbNewLine & "SUMM - " & panel_summ
	Loop until panel_summ = "PNLE"
	' If count = 1 Then ReDim Preserve INCOME_ARRAY(panel_count)
	stop_count = panel_count - 1
	If stop_count < 0 Then stop_count = 0

	' MsgBox start_count & vbNewLine & stop_count
	' If start_count <> 0 Then

	' end If

end function

function read_EATS_panel(hh_membs_eat, unable_to_fix_list, grp_one, grp_one_list, grp_one_array, grp_two, grp_two_list, grp_two_array, grp_three, grp_three_list, grp_three_array, grp_four, grp_four_list, grp_four_array, grp_five, grp_five_list, grp_five_array)
	Call navigate_to_MAXIS_screen("STAT", "EATS")

	EMReadScreen eats_version, 1, 2, 73
	If eats_version = "1" Then
		EMReadScreen hh_membs_eat, 1, 4, 72
		EMReadScreen unable_to_fix_list, 26, 8, 53
		If unable_to_fix_list = "__  __  __  __  __  __  __" Then unable_to_fix_list = ""

		EMReadScreen grp_one, 2, 13, 28
		EMReadScreen grp_one_list, 38, 13, 39
		grp_one_list = replace(grp_one_list, "__", "")
		grp_one_list = trim(grp_one_list)
		grp_one_array = split(grp_one_list, "  ")
		grp_one_list = "MEMB " & replace(grp_one_list, "  ", ", MEMB ")

		EMReadScreen grp_two, 2, 14, 28
		EMReadScreen grp_two_list, 38, 14, 39
		grp_two_list = replace(grp_two_list, "__", "")
		grp_two_list = trim(grp_two_list)
		grp_two_array = split(grp_two_list, "  ")
		grp_two_list = "MEMB " & replace(grp_two_list, "  ", ", MEMB ")

		EMReadScreen grp_three, 2, 15, 28
		EMReadScreen grp_three_list, 38, 15, 39
		grp_three_list = replace(grp_three_list, "__", "")
		grp_three_list = trim(grp_three_list)
		grp_three_array = split(grp_three_list, "  ")
		grp_three_list = "MEMB " & replace(grp_three_list, "  ", ", MEMB ")

		EMReadScreen grp_four, 2, 16, 28
		EMReadScreen grp_four_list, 38, 16, 39
		grp_four_list = replace(grp_four_list, "__", "")
		grp_four_list = trim(grp_four_list)
		grp_four_array = split(grp_four_list, "  ")
		grp_four_list = "MEMB " & replace(grp_four_list, "  ", ", MEMB ")

		EMReadScreen grp_five, 2, 17, 28
		EMReadScreen grp_five_list, 38, 17, 39
		grp_five_list = replace(grp_five_list, "__", "")
		grp_five_list = trim(grp_five_list)
		grp_five_array = split(grp_five_list, "  ")
		grp_five_list = "MEMB " & replace(grp_five_list, "  ", ", MEMB ")
	End If
end function

' FUNCTION - have worker enter the ADDR that is on the CAF or indicate that it is blank - show in dialog and have a droplist for the worker to indicate the client verbally confirmed this address.
' FUNCTION - have worker enter the mailing address on the CAF or indicate that it is blank - show in dialog and have a droplist for the worker to indicate the client verbally confirmaed this address. Add functionality for General Deliver and explaining the requirements
' FUNCTION - list all of the phone numbers from ADDR and have the worker indicate in droplist if they are correct and the type of phone number that it is. If none - then have the worker clarify if there is another phone number or not
function review_ADDR_information()
end function

' FUNCTION - Ask about living situation - worker to select from CAF and verbally confirm. A;sp ask if HOMELESS - if yes then - want to speak to shelter team - do you lack access to work related necessities
function review_living_situation()
end function

' FUNCTION - List all of the HH Members and their relationships. Ask if all of these people still live in the house - Ask if there is any other person living at this address - Confirm the relationships - add new people if needed
function review_the_MEMBs()
end function

' FUNCTION - Lopp through all of the HH Members and confirm:
	' Confirm personal information (name, DOB, marital status, SSN - if not validated)
	' Detail if ID is needed and if we have it
	' Confirm citizenship/immigration
	' Determine language
	' Determine if recently arrived in MN
	' Determine relationship to other HH Members - determine proof if needed
	' Determine who Principal Wage Earner is
function review_MEMB_detail()
end function

' FUNCTION - Get AREP information or confirm no AREP
' FUNCTION - Get SWKR information
function review_associeated_people()
end function


'CAF QUESIION

' 1. Does everyone in your household buy, fix, or eat food with you?
' 2. Is anyone who is in the household, who is age 60 or over or disabled, unable to buy or fix food due to a disability?

'FUNCTION - CAF Q1 & Q2 Enter what CAF has listed and then the verbal confirmation
function reveiw_caf_q_1_2()
end function

' 3. Is anyone in your household attending school?

'FUNCTION - CAF Q3 Enter the CAF answer and the verbal response - if NO move on. If YES then list all members and indicate what kind of school they are attending.
function review_caf_q_3()
end function

' 4. Is anyone in your household temporarily not living in your home? (example: vacation, foster care, treatment, hospital job search)

'FUNCTION - CAF Q4 enter the CAF answer and the verbal response. If No - move on - if yes the who and where they are - explain temp absense
function review_caf_q_4()
end function

' 5.	Is anyone blind, or does anyone have a physical or mental health condition that limit the ability to work or perform daily activities?

' 'FUNCTION - read MAXIS for DISA panels
' function
' end function

'FUNCTION - CAF Q5 enter the CAF answer and the verbal answer. If no - next question, if yes - then WHO and what verifs or program details
function review_caf_q_5()
end function

' 'FUNCTIOn - review the panels and update/delete as needed.
' function
' end function

'6.	Is anyone unable to work for reasons other than illness or disability?

'FUNCTION - CAF Q6 enter CAF answer and verbal response. If no - keep going. If yes - list HH members and the reason for inability to work - this might be important for ABAWD or MFIP extension - EXPAND ON POSSIBLE ABAWD EXEMPTIONS '
function review_caf_q_6()
end function

'7.	In the last 60 days did anyone in the household: Stop working or quit? Refuse a job offer? Ask to work fewwer hours? Go on strike?

'FUNCTION - CAF Q7 enter CAF answer and verbal response. If no - move on. If yes - who, and then details based on program.
function review_caf_q_7()
end function

'8.	Has anyone in the household had a job or been self-employed in the past 12 months?

'FUNCTION - CAF Q8 - enter CAF response and verbal response. If yes gather the details. If SNAP ask about 36 months - CAF response and verbal response - if yes - gather detail'
function review_caf_q_8()
end function

'9.	Does anyone in the household have a job or expect to get income from a job this month or next month? (Include income from Work Study and paid scholarships. Include free benefits or reduced expenses received for work (shelter, food, clothing, etc.)

'FUNCTION - CAF Q9 - enter the CAF response and verbal confirmation. If yes, loop through gathering detail about each job - make sure to gather the pay in app month if at application. Loop until the answer is NO on another job. If SNAP after all jobs ask if anything in the past 36 months.
function review_caf_q_9()
end function

'10. Is anyone in the household self-employed or does anyone expect to get income from self-employment this month or next month?

'FUNCTION - CAF Q10 - enter the CAF response and verbal confirmation. If yes, loop through gathering detail about each self employment - make sure to gather the pay in app month if at application. Loop until the answer is NO on another job. If SNAP after all jobs ask if anything in the past 36 months.
function review_caf_q_10()
end function

'11. Do you expect any changes in income, expenses or work hours?

'FUNCTION - CAF Q11 - enter the CAF response and verbal response - If yes then list all the income and have them enter which is changing, when and how.
function review_caf_q_11()
end function

'12. Has anyone in the household applied for or does anyone get any of the following types of income each month?

'FUNCTION - CAF Q12 - enter the CAF response and verbal confirmation. If yes, loop through gathering detail about each income source or applicaiton - make sure to gather the pay in app month if at application. Loop until the answer is NO on another job. If SNAP after all jobs ask if anything in the past 36 months.
function review_caf_q_12()
end function

'13. Does anyone in the household have or expect to get any loans, scholarships or grants for attending school?

'FUNCTION - CAF Q13 - enter the CAF response and the verbal response. If yes then list the persons and use the school information gathered before to indicate who might have this and what student income there is
function review_caf_q_13()
end function

'14. Does your household have the following housing expenses?
'15. Does your household have the following utility expenses any time during the year?

'FUNCTION - CAF Q14 and Q15 - list all of the expenses with the CAF response and the verbal response then the detail.
function review_caf_q_14_15()
end function

'16. Do you or anyone living with you have costs for care of a child(ren) because you or they are working, looking for work or going to school? The Child Care Assistance Program (CCAP) may help pay child care costs.
'17. Do you or anyone living with you have costs for care of an ill or disabled adult because you are working, looking for work or going to school?
'18. Does anyone in the household pay court-ordered child support, spousal support, child care support, medical support or contribute to a tax-dependent who does not live in your home?

'FUNCTION - CAF Q16 & Q17 & Q18 -  enter the CAF answers and verbal responses. If yes to either - gather details about who, what, how much, and etc
function review_caf_q_16_17_18()
end function

'19. For SNAP only: Does anyone in the household have medical expenses?

'FUNCTION - CAF Q19 - SNAP only - CAF response and verbal response gather detail if Yes
function review_caf_q_19()
end function

'20. Does anyone in the household own, or is anyone buying, any of the following?
'21. FOR CASH PROGRAMS ONLY: Has anyone in the household given away, sold or traded anything of value in the past 12 months? (For Example: Cash, Bank Accounts, Stocks, Bonds, or Vehicles)?

'FUNCTION - CAF Q20 & Q21 - All the asset information - for cash programs we need some verif - for SNAP if at application - enter balances - no verif - EXP
function review_caf_q_20_21()
end function

'22. FOR RECERTIFICATIONS ONLY: Did anyone move in or out of your home in the past 12 months?

'FUNCTION - CAF Q22 - Enter CAF response and verbal response
function review_caf_q_22()
end function

'23. For children under the age of 19, are both parents living in the home?

'FUNCTION - CAF Q23 - Enter CAF response and verbal response
function review_caf_q_23()
end function

'24. FOR MSA RECIPIENTS ONLY: Does anyone in the household have any of the following expenses?

'FUNCTION - CAF Q24 - Enter CAF response and verbal response
function review_caf_q_24()
end function

'25. Penalty warning and qualification questions

'FUNCTION - CAF Q25 - Enter CAF response and verbal response
function review_caf_q_25()
end function

'26. Did client sign the last page of the CAF?

'FUNCTION - CAF Q26 - Enter CAF response and verbal response
function review_caf_q_26()
end function

Dim household_homeless_response

'THE SCRIPT ===========================================================================================================
EMConnect ""
Call check_for_MAXIS(TRUE)

Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
interview_date = date & ""
case_has_income_listed = FALSE

Call back_to_SELF
EMReadScreen MX_region, 10, 22, 48
MX_region = trim(MX_region)
If MX_region = "INQUIRY DB" Then
	continue_in_inquiry = MsgBox("You have started this script run in INQUIRY." & vbNewLine & vbNewLine & "The script cannot complete a CASE:NOTE when run in inquiry. The functionality is limited when run in inquiry. " & vbNewLine & vbNewLine & "Would you like to continue in INQUIRY?", vbQuestion + vbYesNo, "Continue in INQUIRY")
	If continue_in_inquiry = vbNo Then Call script_end_procedure("~PT Interview Script cancelled as it was run in inquiry.")
End If
'
'Start of Interview
'Dialog to gather interview set up information/detail.
'Date of interview, type of interview, case number.
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 281, 105, "INTERVIEW Case number dialog"
  EditBox 65, 10, 60, 15, MAXIS_case_number
  EditBox 210, 10, 15, 15, MAXIS_footer_month
  EditBox 230, 10, 15, 15, MAXIS_footer_year
  CheckBox 10, 45, 30, 10, "CASH", CASH_on_CAF_checkbox
  CheckBox 50, 45, 35, 10, "SNAP", SNAP_on_CAF_checkbox
  CheckBox 90, 45, 35, 10, "EMER", EMER_on_CAF_checkbox
  DropListBox 135, 45, 140, 15, "Select One:"+chr(9)+"CAF (DHS-5223)"+chr(9)+"SNAP App for Srs (DHS-5223F)"+chr(9)+"ApplyMN"+chr(9)+"Combined AR for Certain Pops (DHS-3727)"+chr(9)+"CAF Addendum (DHS-5223C)", CAF_form
  EditBox 60, 65, 45, 15, interview_date
  ComboBox 165, 65, 110, 45, "Type or Select"+chr(9)+"phone"+chr(9)+"office", interview_type
  ButtonGroup ButtonPressed
    PushButton 35, 85, 15, 15, "!", tips_and_tricks_button
    OkButton 170, 85, 50, 15
    CancelButton 225, 85, 50, 15
  GroupBox 5, 30, 125, 30, "Programs marked on CAF"
  Text 135, 35, 65, 10, "Actual CAF Form:"
  Text 55, 90, 105, 10, "Look for me for Tips and Tricks!"
  Text 10, 70, 50, 10, "Interview Date:"
  Text 115, 70, 50, 10, "Interview Type:"
  Text 140, 15, 65, 10, "Footer month/year:"
  Text 10, 15, 50, 10, "Case number:"
EndDialog

Do
	Do
		err_msg = ""

		dialog Dialog1
		cancel_without_confirmation

		If err_msg <> "" Then MsgBox "*** Please resolve to Continue: ***" & vbNewLine & err_msg

	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE

Call read_ADDR_panel(addr_eff_date, resi_line_one, resi_line_two, resi_city, resi_state, resi_zip, resi_county, verif, homeless, ind_reservation, living_sit, res_name, mail_line_one, mail_line_two, mail_city, mail_state, mail_zip, phone_numb_one, phone_type_one, phone_numb_two, phone_type_two, phone_numb_three, phone_type_three, addr_updated_date)
updated_addr_eff_date 		= addr_eff_date
updated_resi_line_one 		= resi_line_one
updated_resi_line_two 		= resi_line_two
updated_resi_city 			= resi_city
updated_resi_state 			= resi_state
updated_resi_zip 			= resi_zip
updated_resi_county 		= resi_county
updated_verif 				= verif
updated_homeless 			= homeless
updated_ind_reservation 	= ind_reservation
updated_living_sit 			= living_sit
updated_res_name 			= res_name
updated_mail_line_one 		= mail_line_one
updated_mail_line_two 		= mail_line_two
updated_mail_city 			= mail_city
updated_mail_state 			= mail_state
updated_mail_zip 			= mail_zip
updated_phone_numb_one 		= phone_numb_one
updated_phone_type_one 		= phone_type_one
updated_phone_numb_two 		= phone_numb_two
updated_phone_type_two 		= phone_type_two
updated_phone_numb_three 	= phone_numb_three
updated_phone_type_three 	= phone_type_three

Call read_all_the_MEMBs

Call read_all_Assets

Call read_all_UNEA


call read_EATS_panel(all_hh_members_eat_w_applicant, members_unable_to_fix_food, group_one_number, group_one_member_list, group_one_member_array, group_two_number, group_two_member_list, group_two_member_array, group_three_number, group_three_member_list, group_three_member_array, group_four_number, group_four_member_list, group_four_member_array, group_five_number, group_five_member_list, group_five_member_array)


For hh_memb = 0 to UBound(HH_MEMB_ARRAY, 1)
	' MsgBox HH_MEMB_ARRAY(hh_memb).ref_number & vbNewLine & HH_MEMB_ARRAY(hh_memb).full_name
	If case_has_income_listed = TRUE Then
		For each_income = 0 to UBound(INCOME_ARRAY, 1)
			' MsgBox INCOME_ARRAY(each_income)
			' MsgBox "1 - " & INCOME_ARRAY(each_income).member_ref
			' MsgBox "2 - " & HH_MEMB_ARRAY(hh_memb).ref_number

			If INCOME_ARRAY(each_income).member_ref = HH_MEMB_ARRAY(hh_memb).ref_number Then
				If INCOME_ARRAY(each_income).panel_name = "JOBS" Then HH_MEMB_ARRAY(hh_memb).clt_has_JOBS = TRUE
				If INCOME_ARRAY(each_income).panel_name = "BUSI" Then HH_MEMB_ARRAY(hh_memb).clt_has_BUSI = TRUE
				If INCOME_ARRAY(each_income).panel_name = "UNEA" Then
					If INCOME_ARRAY(i).income_type_code = "08" OR INCOME_ARRAY(i).income_type_code = "36" OR INCOME_ARRAY(i).income_type_code = "39" OR INCOME_ARRAY(i).income_type_code = "43" OR INCOME_ARRAY(i).income_type_code = "45" Then HH_MEMB_ARRAY(hh_memb).clt_has_cs_income = TRUE
					If INCOME_ARRAY(i).income_type_code = "35" OR INCOME_ARRAY(i).income_type_code = "37" OR INCOME_ARRAY(i).income_type_code = "40" Then HH_MEMB_ARRAY(hh_memb).clt_has_ss_income = TRUE
				End If
			End If
		Next
	End If
Next
'Gather all of the household information

'Go to see if recert or application and CAF date
'Gather all of the HH Member information
'Gather the DISA information

'Dialog to get the programs, person the interview is completed with, and some basic CAF information
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 326, 45, "Interview Details"
  EditBox 50, 5, 50, 15, caf_date
  ComboBox 175, 5, 145, 45, "Type or Select"+chr(9)+"List of People", interview_with
  ButtonGroup ButtonPressed
    OkButton 220, 25, 50, 15
    CancelButton 270, 25, 50, 15
  Text 10, 10, 35, 10, "CAF Date:"
  Text 120, 10, 50, 10, "Interview With:"
EndDialog
Do
	Do
		err_msg = ""

		dialog Dialog1
		cancel_without_confirmation

		If err_msg <> "" Then MsgBox "*** Please resolve to Continue: ***" & vbNewLine & err_msg

	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE


page_display = show_pg_one
second_page_display = main_unea

ButtonPressed = caf_page_one_btn
leave_loop = FALSE

If unea_found = FALSE AND jobs_found = FALSE AND busi_found = FALSE Then
	Set INCOME_ARRAY(0) = new client_income
End If
Do
	Do
		If memb_selected = "" Then memb_selected = 0

		case_has_stwk_panel = FALSE
		If page_display = show_q_7 AND stwk_selected = "" Then
			For i = 0 to Ubound(HH_MEMB_ARRAY, 1)
				If HH_MEMB_ARRAY(i).stwk_exists = TRUE Then
					stwk_selected = i
					case_has_stwk_panel = TRUE
					Exit For
				End If
			Next
		End If
		case_has_fmed_panel = FALSE
		If page_display = show_q_19 AND fmed_selected = "" Then
			For i = 0 to Ubound(HH_MEMB_ARRAY, 1)
				If HH_MEMB_ARRAY(i).fmed_exists = TRUE Then
					fmed_selected = i
					case_has_fmed_panel = TRUE
					Exit For
				End If
			Next
		End If
		' MsgBox page_display
		Dialog1 = ""
		call define_main_dialog

		err_msg = ""

		prev_page = page_display


		dialog Dialog1
		cancel_confirmation

		If err_msg <> "" Then MsgBox "*** Please resolve to Continue: ***" & vbNewLine & err_msg

		If page_display <> prev_page Then
			'ADD FUNCTIONS HERE TO EVALUATE THE COMPLETION OF EACH PAGE
		End If

		' MsgBox "ButtonPressed - " & ButtonPressed

		call dialog_movement


	Loop until leave_loop = TRUE
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE


Do
	Do
		err_msg = ""

		BeginDialog Dialog1, 0, 0, 550, 385, "Full Interview Questions"
		  ButtonGroup ButtonPressed
		    PushButton 465, 365, 80, 15, "Review Office Contact", move_to_contact
		  Text 10, 10, 160, 10, "^^1 - READ the complete text here to the client:"
		  GroupBox 10, 25, 530, 325, "Rights and Responsibilities Text"
		  Text 20, 35, 505, 25, "Note: Cash on an Electronic Benefit Transfer (EBT) card is provided to help families meet their basic needs, including: food, shelter, clothing, utilities and transportation. These funds are provided until families can support themselves. It is illegal for an EBT user to buy or attempt to buy tobacco products or alcohol with the EBT card. If you do, it is fraud and you will be removed from the program. Do not use an EBT card at a gambling establishment or "
		  Text 20, 60, 490, 10, "retail establishment, which provides adult-orientated entertainment in which performers disrobe or perform in an unclothed state for entertainment."
		  Text 20, 80, 215, 155, "- If you receive cash assistance and/or child care assistance, you must report changes which may affect your benefits to the county agency within 10 days after the change has occurred. If you receive"
		EndDialog

		dialog Dialog1
	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = FALSE


'CASE:NOTES
'SEPERATE notes with pertinent information. This will make it easier to read
'Have a title convention that will make it clear these are all together (indents or maybe same start of header)' Idea on training case 298532

'CAF PAGE 1 - HH Comp and Address
' FUNCTION - Read the current address on ADDR - resi and mail and phone numbers.
' FUNCTION - have worker enter the ADDR that is on the CAF or indicate that it is blank - show in dialog and have a droplist for the worker to indicate the client verbally confirmed this address.
' FUNCTION - have worker enter the mailing address on the CAF or indicate that it is blank - show in dialog and have a droplist for the worker to indicate the client verbally confirmaed this address. Add functionality for General Deliver and explaining the requirements
' FUNCTION - list all of the phone numbers from ADDR and have the worker indicate in droplist if they are correct and the type of phone number that it is. If none - then have the worker clarify if there is another phone number or not

' FUNCTION - Ask about living situation - worker to select from CAF and verbally confirm. A;sp ask if HOMELESS - if yes then - want to speak to shelter team - do you lack access to work related necessities
' FUNCTION - Read all the HH Members from case, including how they are related to M01 and SIBL and PARE - get age
' FUNCTION - List all of the HH Members and their relationships. Ask if all of these people still live in the house - Ask if there is any other person living at this address - Confirm the relationships
' FUNCTION - Lopp through all of the HH Members and confirm:
	' Confirm personal information (name, DOB, marital status, SSN - if not validated)
	' Detail if ID is needed and if we have it
	' Confirm citizenship/immigration
	' Determine language
	' Determine if recently arrived in MN
	' Determine relationship to other HH Members - determine proof if needed
	' Determine who Principal Wage Earner is
' FUNCTION - Get AREP information or confirm no AREP
' FUNCTION - Get SWKR information


'CAF QUESIION

' 1. Does everyone in your household buy, fix, or eat food with you?
' 2. Is anyone who is in the household, who is age 60 or over or disabled, unable to buy or fix food due to a disability?

'FUNCTION - CAF Q1 & Q2 Enter what CAF has listed and then the verbal confirmation

' 3. Is anyone in your household attending school?

'FUNCTION - CAF Q3 Enter the CAF answer and the verbal response - if NO move on. If YES then list all members and indicate what kind of school they are attending.

' 4. Is anyone in your household temporarily not living in your home? (example: vacation, foster care, treatment, hospital job search)

'FUNCTION - CAF Q4 enter the CAF answer and the verbal response. If No - move on - if yes the who and where they are - explain temp absense

' 5.	Is anyone blind, or does anyone have a physical or mental health condition that limit the ability to work or perform daily activities?

'FUNCTION - read MAXIS for DISA panels
'FUNCTION - CAF Q5 enter the CAF answer and the verbal answer. If no - next question, if yes - then WHO and what verifs or program details
'FUNCTIOn - review the panels and update/delete as needed.

'6.	Is anyone unable to work for reasons other than illness or disability?

'FUNCTION - CAF Q6 enter CAF answer and verbal response. If no - keep going. If yes - list HH members and the reason for inability to work - this might be important for ABAWD or MFIP extension - EXPAND ON POSSIBLE ABAWD EXEMPTIONS '

'7.	In the last 60 days did anyone in the household: Stop working or quit? Refuse a job offer? Ask to work fewwer hours? Go on strike?

'FUNCTION - CAF Q7 enter CAF answer and verbal response. If no - move on. If yes - who, and then details based on program.

'8.	Has anyone in the household had a job or been self-employed in the past 12 months?

'FUNCTION - CAF Q8 - enter CAF response and verbal response. If yes gather the details. If SNAP ask about 36 months - CAF response and verbal response - if yes - gather detail'

'9.	Does anyone in the household have a job or expect to get income from a job this month or next month? (Include income from Work Study and paid scholarships. Include free benefits or reduced expenses received for work (shelter, food, clothing, etc.)

'FUNCTION - CAF Q9 - enter the CAF response and verbal confirmation. If yes, loop through gathering detail about each job - make sure to gather the pay in app month if at application. Loop until the answer is NO on another job. If SNAP after all jobs ask if anything in the past 36 months.

'10. Is anyone in the household self-employed or does anyone expect to get income from self-employment this month or next month?

'FUNCTION - CAF Q10 - enter the CAF response and verbal confirmation. If yes, loop through gathering detail about each self employment - make sure to gather the pay in app month if at application. Loop until the answer is NO on another job. If SNAP after all jobs ask if anything in the past 36 months.

'11. Do you expect any changes in income, expenses or work hours?

'FUNCTION - CAF Q11 - enter the CAF response and verbal response - If yes then list all the income and have them enter which is changing, when and how.

'12. Has anyone in the household applied for or does anyone get any of the following types of income each month?

'FUNCTION - CAF Q12 - enter the CAF response and verbal confirmation. If yes, loop through gathering detail about each income source or applicaiton - make sure to gather the pay in app month if at application. Loop until the answer is NO on another job. If SNAP after all jobs ask if anything in the past 36 months.

'13. Does anyone in the household have or expect to get any loans, scholarships or grants for attending school?

'FUNCTION - CAF Q13 - enter the CAF response and the verbal response. If yes then list the persons and use the school information gathered before to indicate who might have this and what student income there is

'14. Does your household have the following housing expenses?
'15. Does your household have the following utility expenses any time during the year?

'FUNCTION - CAF Q14 and Q15 - list all of the expenses with the CAF response and the verbal response then the detail.

'16. Do you or anyone living with you have costs for care of a child(ren) because you or they are working, looking for work or going to school? The Child Care Assistance Program (CCAP) may help pay child care costs.
'17. Do you or anyone living with you have costs for care of an ill or disabled adult because you are working, looking for work or going to school?
'18. Does anyone in the household pay court-ordered child support, spousal support, child care support, medical support or contribute to a tax-dependent who does not live in your home?

'FUNCTION - CAF Q16 & Q17 & Q18 -  enter the CAF answers and verbal responses. If yes to either - gather details about who, what, how much, and etc

'19. For SNAP only: Does anyone in the household have medical expenses?

'FUNCTION - CAF Q19 - SNAP only - CAF response and verbal response gather detail if Yes

'20. Does anyone in the household own, or is anyone buying, any of the following?
'21. FOR CASH PROGRAMS ONLY: Has anyone in the household given away, sold or traded anything of value in the past 12 months? (For Example: Cash, Bank Accounts, Stocks, Bonds, or Vehicles)?

'FUNCTION - CAF Q20 & Q21 - All the asset information - for cash programs we need some verif - for SNAP if at application - enter balances - no verif - EXP

'22. FOR RECERTIFICATIONS ONLY: Did anyone move in or out of your home in the past 12 months?

'FUNCTION - CAF Q22 - Enter CAF response and verbal response

'23. For children under the age of 19, are both parents living in the home?

'FUNCTION - CAF Q23 - Enter CAF response and verbal response

'24. FOR MSA RECIPIENTS ONLY: Does anyone in the household have any of the following expenses?

'FUNCTION - CAF Q24 - Enter CAF response and verbal response

'25. Penalty warning and qualification questions

'FUNCTION - CAF Q25 - Enter CAF response and verbal response

'26. Did client sign the last page of the CAF?

'FUNCTION - CAF Q26 - Enter CAF response and verbal response
