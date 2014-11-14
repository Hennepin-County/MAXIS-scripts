FUNCTION write_panel_to_maxis_UNEA(unea_inc_type, unea_inc_verif, unea_claim_suffix, unea_start_date, unea_pay_freq, unea_inc_amount, ssn_first, ssn_mid, ssn_last)
	call navigate_to_screen("STAT", "UNEA")
	EMWriteScreen reference_number, 20, 76
	EMWriteScreen "NN", 20, 79
	transmit

	EMWriteScreen unea_inc_type, 5, 37
	EMWriteScreen unea_inc_verif, 5, 65
	EMWriteScreen (ssn_first & ssn_mid & ssn_last & unea_claim_suffix), 6, 37
	call create_maxis_friendly_date(unea_start_date, 0, 7, 37)

	'=====Navigates to the PIC for UNEA=====
	EMWriteScreen "X", 10, 26
	transmit
	EMWriteScreen unea_pay_freq, 5, 64
	EMWriteScreen unea_inc_amount, 8, 66
	calc_month = datepart("M", date)
		IF len(calc_month) = 1 THEN calc_month = "0" & calc_month
	calc_day = datepart("D", date)
		IF len(calc_day) = 1 THEN calc_day = "0" & calc_day
	calc_year = datepart("YYYY", date)
	EMWriteScreen calc_month, 5, 34
	EMWriteScreen calc_day, 5, 37
	EMWriteScreen calc_year, 5, 40
	transmit
	transmit
	transmit		'<=====navigates out of the PIC

	'=====the following bit is for the retrospective & prospective pay dates=====
	EMReadScreen bene_month, 2, 20, 55
	EMReadScreen bene_year, 2, 20, 58
	retro_month = bene_month - 2
	retro_year = bene_year
		IF retro_month < 1 THEN
			retro_month = bene_month + 10
			retro_year = bene_year - 1
		END IF

	EMWriteScreen retro_month, 13, 25
	EMWriteScreen "05", 13, 28
	EMWriteScreen retro_year, 13, 31
	EMWriteScreen "________", 13, 39
	EMWriteScreen unea_inc_amount, 13, 39
	EMWriteScreen bene_month, 13, 54
	EMWriteScreen "05", 13, 57
	EMWriteScreen bene_year, 13, 60
	EMWriteScreen "________", 13, 68
	EMWriteScreen unea_inc_amount, 13, 68
	
	IF unea_pay_freq = "2" OR unea_pay_freq = "3" THEN
		EMWriteScreen retro_month, 14, 25
		EMWriteScreen "19", 14, 28
		EMWriteScreen retro_year, 14, 31
		EMWriteScreen "________", 14, 39
		EMWriteScreen unea_inc_amount, 14, 39
		EMWriteScreen bene_month, 14, 54
		EMWriteScreen "19", 14, 57
		EMWriteScreen bene_year, 14, 60
		EMWriteScreen "________", 14, 68
		EMWriteScreen unea_inc_amount, 14, 68
	ELSEIF unea_pay_freq = "4" THEN
		EMWriteScreen retro_month, 14, 25
		EMWriteScreen "12", 14, 28
		EMWriteScreen retro_year, 14, 31
		EMWriteScreen "________", 14, 39
		EMWriteScreen unea_inc_amount, 14, 39
		EMWriteScreen retro_month, 15, 25
		EMWriteScreen "19", 15, 28
		EMWriteScreen retro_year, 15, 31
		EMWriteScreen "________", 15, 39
		EMWriteScreen unea_inc_amount, 15, 39
		EMWriteScreen retro_month, 16, 25
		EMWriteScreen "26", 16, 28
		EMWriteScreen retro_year, 16, 31
		EMWriteScreen "________", 16, 39
		EMWriteScreen unea_inc_amount, 16, 39
		EMWriteScreen bene_month, 14, 54
		EMWriteScreen "12", 14, 57
		EMWriteScreen bene_year, 14, 60
		EMWriteScreen "________", 14, 68
		EMWriteScreen unea_inc_amount, 14, 68
		EMWriteScreen bene_month, 15, 54 
		EMWriteScreen "19", 15, 57 
		EMWriteScreen bene_year, 15, 60 
		EMWriteScreen "________", 15, 68 
		EMWriteScreen unea_inc_amount, 15, 68 
		EMWriteScreen bene_month, 16, 54
		EMWriteScreen "26", 16, 57
		EMWriteScreen bene_year, 16, 60
		EMWriteScreen "________", 16, 68
		EMWriteScreen unea_inc_amount, 16, 68
	END IF

	'=====determines if the benefit month is current month + 1 and dumps information into the HC income estimator
	IF (bene_month * 1) = (datepart("M", date) + 1) THEN		'<===== "bene_month * 1" is needed to convert bene_month from a string to a useable number
		EMWriteScreen "X", 6, 56
		transmit
		EMWriteScreen "________", 9, 65
		EMWriteScreen unea_inc_amount, 9, 65
		EMWriteScreen unea_pay_freq, 10, 63
		transmit
		transmit
	END IF

END FUNCTION


FUNCTION write_panel_to_maxis_DIET(mfip_diet1, mfip_dietv1, mfip_diet2, mfip_dietv2, msa_diet1, msa_dietv1, msa_diet2, msa_dietv2, msa_diet3, msa_dietv3, msa_diet4, msa_dietv4)
	IF mfip_diet1 <> "" AND mfip_diet2 <> "" AND msa_diet1 <> "" AND msa_diet2 <> "" AND msa_diet3 <> "" AND msa_diet4 <> "" THEN
		call navigate_to_screen("STAT", "DIET")
		EMWriteScreen reference_number, 20, 76
		EMWriteScreen "NN", 20, 79
		transmit

		EMWriteScreen mfip1, 8, 40
		EMWriteScreen mfipv1, 8, 51
		EMWriteScreen mfip2, 9, 40
		EMWriteScreen mfipv2, 9, 51
		EMWriteScreen msa1, 11, 40
		EMWriteScreen msav1, 11, 51
		EMWriteScreen msa2, 12, 40
		EMWriteScreen msav2, 12, 51
		EMWriteScreen msa3, 13, 40
		EMWriteScreen msav3, 13, 51
		EMWriteScreen msa4, 14, 40
		EMWriteScreen msav4, 14, 51
		transmit
END FUNCTION


FUNCTION write_panel_to_maxis_MMSA(mmsa_liv_arr, mmsa_cont_elig, mmsa_spous_inc, mmsa_shared_hous)
	IF mmsa_liv_arr <> "" THEN
		call navigate_to_screen("STAT", "MMSA")
		EMWriteScreen "NN", 20, 79
		transmit
		EMWriteScreen mmsa_liv_arr, 7, 54
		EMWriteScreen mmsa_cont_elig, 9, 54
		EMWriteScreen mmsa_spous_inc, 12, 62
		EMWriteScreen mmsa_shared_hous, 14, 62
		transmit
	END IF
END FUNCTION


FUNCTION write_panel_to_maxis_EATS(eats_together, eats_boarder, eats_group_one, eats_group_two, eats_group_three)
	IF reference_number = "01" THEN
		call navigate_to_screen("STAT", "EATS")
		EMWriteScreen eats_together, 4, 72
		EMWriteScreen eats_boarder, 5, 72
		IF ucase(eats_together) = "N" THEN
			EMWriteScreen "01", 13, 28
			eats_group_one = replace(eats_group_one, " ", "")
			eats_group_one = split(eats_group_one, ",")
			eats_col = 39
			FOR EACH eats_household_member IN eats_group_one
				EMWriteScreen eats_household_member, 13, eats_col
				eats_col = eats_col + 4
			NEXT
			EMWriteScreen "02", 14, 28
			eats_group_two = replace(eats_group_two, " ", "")
			eats_group_two = split(eats_group_two, ",")
			eats_col = 39
			FOR EACH eats_household_member IN eats_group_two
				EMWriteScreen eats_household_member, 14, eats_col
				eats_col = eats_col + 4
			NEXT
			IF eats_group_three <> "" THEN
				EMWriteScreen "03", 15, 28
				eats_group_three = replace(eats_group_three, " ", "")
				eats_group_three = split(eats_group_three, ",")
				eats_col = 39
				FOR EACH eats_household_member IN eats_group_three
					EMWriteScreen eats_household_member, 15, eats_col
					eats_col = eats_col + 4
				NEXT
			END IF
		END IF
	transmit
	END IF
END FUNCTION


FUNCTION write_panel_to_maxis_WREG(wreg_fs_pwe, wreg_fset_status, wreg_defer_fs, wreg_fset_orientation_date, wreg_fset_sanction_date, wreg_num_sanctions, wreg_abawd_status, wreg_ga_basis)
	call navigate_to_screen("STAT", "WREG")
	EMWriteScreen reference_number, 20, 76
	EMWriteScreen "NN", 20, 79
	transmit

	EMWriteScreen wreg_fs_pwe, 6, 68
	EMWriteScreen wreg_fset_status, 8, 50
	EMWriteScreen wreg_defer_fs, 8, 80
	call create_maxis_friendly_date(wreg_fset_orientation_date, 0, 9, 50)
	IF wreg_fset_sanction_date <> "" THEN call create_maxis_friendly_date(wreg_fset_orientation_date, 0, 10, 50)
	IF wreg_num_sanctions <> "" THEN EMWriteScreen wreg_num_sanctions, 11, 50
	EMWriteScreen wreg_abawd_status, 13, 50
	EMWriteScreen wreg_ga_basis, 15, 50

	transmit
END FUNCTION


FUNCTION write_panel_to_maxis_TYPE_PROG_REVW(appl_date, type_cash_yn, type_hc_yn, type_fs_yn, prog_mig_worker, revw_ar_or_ir, revw_exempt)
	call navigate_to_screen("STAT", "TYPE")
	IF reference_number = "01" THEN
		EMWriteScreen "NN", 20, 79
		transmit
		EMWriteScreen type_cash_yn, 6, 28
		EMWriteScreen type_hc_yn, 6, 37
		EMWriteScreen type_fs_yn, 6, 46
		EMWriteScreen "N", 6, 55
		EMWriteScreen "N", 6, 64
		EMWriteScreen "N", 6, 73
		type_row = 7
		DO				'<=====this DO/LOOP populates "N" for all other HH members on TYPE so the script can get past TYPE when the reference number = "01"
			EMReadScreen type_does_hh_memb_exist, 2, type_row, 3
			IF type_does_hh_memb_exist <> "  " THEN
				EMWriteScreen "N", type_row, 28
				EMWriteScreen "N", type_row, 37
				EMWriteScreen "N", type_row, 46
				EMWriteScreen "N", type_row, 55
				type_row = type_row + 1
			ELSE
				EXIT DO
			END IF
		LOOP WHILE type_does_hh_memb_exist <> "  "
	ELSE
		PF9
		type_row = 7
		DO
			EMReadScreen type_does_hh_memb_exist, 2, type_row, 3
			IF type_does_hh_memb_exist = reference_number THEN
				EMWriteScreen type_cash_yn, type_row, 28
				EMWriteScreen type_hc_yn, type_row, 37
				EMWriteScreen type_fs_yn, type_row, 46
				EMWriteScreen "N", type_row, 55
				exit do
			ELSE
				type_row = type_row + 1
			END IF
		LOOP UNTIL type_does_hh_memb_exist = reference_number
	END IF	
	transmit		'<===== when reference_number = "01" this transmit will navigate to PROG, else, it will navigate to STAT/WRAP

	IF reference_number = "01" THEN		'<===== only accesses PROG & REVW if reference_number = "01"
		call navigate_to_screen("STAT", "PROG")
		EMWriteScreen "NN", 20, 71
		transmit
			IF type_cash_yn = "Y" THEN
				call create_maxis_friendly_date(appl_date, 0, 6, 33)
				call create_maxis_friendly_date(appl_date, 0, 6, 44)
				call create_maxis_friendly_date(appl_date, 0, 6, 55)
			END IF
			IF type_fs_yn = "Y" THEN
				call create_maxis_friendly_date(appl_date, 0, 10, 33)
				call create_maxis_friendly_date(appl_date, 0, 10, 44)
				call create_maxis_friendly_date(appl_date, 0, 10, 55)
			END IF
			IF type_hc_yn = "Y" THEN
				call create_maxis_friendly_date(appl_date, 0, 12, 33)
				call create_maxis_friendly_date(appl_date, 0, 12, 55)
			END IF
			EMWriteScreen mig_worker, 18, 67
			transmit
			EMWriteScreen mig_worker, 18, 67
			transmit

		call navigate_to_screen("STAT", "REVW")
		EMWriteScreen "NN", 20, 71
		transmit
			IF type_cash_yn = "Y" THEN
				cash_review_date = dateadd("YYYY", 1, appl_date)
				call create_maxis_friendly_date(cash_review_date, 0, 9, 37)
			END IF
			IF type_fs_yn = "Y" THEN
				EMWriteScreen "X", 5, 58
				transmit
				DO
					EMReadScreen food_support_reports, 20, 5, 30
				LOOP UNTIL food_support_reports = "FOOD SUPPORT REPORTS"
				fs_csr_date = dateadd("M", 6, appl_date)
				fs_er_date = dateadd("M", 12, appl_date)
				call create_maxis_friendly_date(fs_csr_date, 0, 9, 26)
				call create_maxis_friendly_date(fs_er_date, 0 9, 64)
				transmit
			END IF
			IF type_hc_yn = "Y" THEN
				EMWriteScreen "X", 5, 71
				transmit
				DO
					EMReadScreen health_care_renewals, 20, 4, 32
				LOOP UNTIL health_care_renewals = "HEALTH CARE RENEWALS"
				IF revw_ar_or_ir = "AR" THEN
					call create_maxis_friendly_date((dateadd("M", 6, appl_date)), 0, 8, 71)
				ELSEIF revw_ar_or_ir = "IR" THEN
					call create_maxis_friendly_date((dateadd("M", 6, appl_date)), 0, 8, 27)
				END IF
				call create_maxis_friendly_date((dateadd("M", 12, appl_date)), 0, 9, 27)
				EMWriteScreen revw_exempt, 9, 71
				transmit
			END IF
	END IF
END FUNCTION


FUNCTION write_panel_to_maxis_JOBS(jobs_inc_type, jobs_inc_verif, jobs_employer_name, jobs_inc_start, jobs_wkly_hrs, jobs_hrly_wage, jobs_pay_freq)
	call navigate_to_screen("STAT", "JOBS")
	EMWriteScreen reference_number, 20, 76
	EMWriteScreen "NN", 20, 79
	transmit

	EMWriteScreen jobs_inc_type, 5, 38
	EMWriteScreen jobs_inc_verif, 6, 38
	EMWriteScreen jobs_employer_name, 7, 42
	call create_maxis_friendly_date(jobs_inc_start, 0, 9, 35)
	EMWriteScreen jobs_pay_freq, 18, 35
	
	'===== navigates to the SNAP PIC to update the PIC =====
	EMWriteScreen "X", 19, 38
	transmit
	DO
		EMReadScreen at_snap_pic, 12, 3, 22
	LOOP UNTIL at_snap_pic = "Food Support"
	call create_maxis_friendly_date(date, 0, 5, 34)
	EMWriteScreen jobs_pay_freq, 5, 64
	EMWriteScreen jobs_wkly_hrs, 8, 64
	EMWriteScreen jobs_hrly_wage, 9, 66
	transmit
	transmit
	EMReadScreen jobs_pic_hrs_per_pp, 6, 16, 51
	EMReadScreen jobs_pic_wages_per_pp, 7, 17, 57
	transmit		'<=====navigates out of the PIC

		'=====the following bit is for the retrospective & prospective pay dates=====
	EMReadScreen bene_month, 2, 20, 55
	EMReadScreen bene_year, 2, 20, 58
	retro_month = bene_month - 2
	retro_year = bene_year
		IF retro_month < 1 THEN
			retro_month = bene_month + 10
			retro_year = bene_year - 1
		END IF

	EMWriteScreen retro_month, 12, 25
	EMWriteScreen "05", 12, 28
	EMWriteScreen retro_year, 12, 31
	EMWriteScreen "________", 12, 38
	EMWriteScreen jobs_pic_wages_per_pp, 12, 38
	EMWriteScreen bene_month, 12, 54
	EMWriteScreen "05", 12, 57
	EMWriteScreen bene_year, 12, 60
	EMWriteScreen "________", 12, 67
	EMWriteScreen jobs_pic_wages_per_pp, 12, 67
	
	IF jobs_pay_freq = "2" OR jobs_pay_freq = "3" THEN
		EMWriteScreen retro_month, 13, 25
		EMWriteScreen "19", 13, 28
		EMWriteScreen retro_year, 13, 31
		EMWriteScreen "________", 13, 38
		EMWriteScreen jobs_pic_wages_per_pp, 13, 38
		EMWriteScreen bene_month, 13, 54
		EMWriteScreen "19", 13, 57
		EMWriteScreen bene_year, 13, 60
		EMWriteScreen "________", 13, 67
		EMWriteScreen jobs_pic_wages_per_pp, 13, 67
	ELSEIF pay_freq = "4" THEN
		EMWriteScreen retro_month, 13, 25
		EMWriteScreen "12", 13, 28
		EMWriteScreen retro_year, 13, 31
		EMWriteScreen "________", 13, 38
		EMWriteScreen jobs_pic_wages_per_pp, 13, 38
		EMWriteScreen retro_month, 14, 25
		EMWriteScreen "19", 14, 28
		EMWriteScreen retro_year, 14, 31
		EMWriteScreen "________", 14, 38
		EMWriteScreen jobs_pic_wages_per_pp, 14, 38
		EMWriteScreen retro_month, 15, 25
		EMWriteScreen "26", 15, 28
		EMWriteScreen retro_year, 15, 31
		EMWriteScreen "________", 15, 38
		EMWriteScreen jobs_pic_wages_per_pp, 15, 38
		EMWriteScreen bene_month, 13, 54
		EMWriteScreen "12", 13, 57
		EMWriteScreen bene_year, 13, 60
		EMWriteScreen "________", 13, 67
		EMWriteScreen jobs_pic_wages_per_pp, 13, 67
		EMWriteScreen bene_month, 14, 54 
		EMWriteScreen "19", 14, 57 
		EMWriteScreen bene_year, 14, 60 
		EMWriteScreen "________", 14, 67 
		EMWriteScreen jobs_pic_wages_per_pp, 14, 67
		EMWriteScreen bene_month, 15, 54
		EMWriteScreen "26", 15, 57
		EMWriteScreen bene_year, 15, 60
		EMWriteScreen "________", 15, 67
		EMWriteScreen jobs_pic_wages_per_pp, 15, 67
	END IF

	'=====determines if the benefit month is current month + 1 and dumps information into the HC income estimator
	IF (bene_month * 1) = (datepart("M", DATE) + 1) THEN		'<===== "bene_month * 1" is needed to convert bene_month from a string to a useable number
		EMWriteScreen "X", 19, 54
		transmit
		EMWriteScreen "________", 9, 65
		EMWriteScreen jobs_pic_wages_per_pp, 11, 63
		transmit
		transmit
	END IF
END FUNCTION
