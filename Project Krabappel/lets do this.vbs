FUNCTION write_panel_to_maxis_UNEA(inc_type, inc_verif, claim_num, start_date, pay_freq, inc_amount)
	EMWriteScreen inc_type, 5, 37
	EMWriteScreen inc_verif, 5, 65
	EMWriteScreen claim_num, 6, 37
	call create_maxis_friendly_date(start_date, 0, 7, 37)

	'=====Navigates to the PIC for UNEA=====
	EMWriteScreen "X", 10, 26
	transmit
	EMWriteScreen pay_freq, 5, 64
	EMWriteScreen inc_amount, 8, 66
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
	EMWriteScreen inc_amount, 13, 39
	EMWriteScreen bene_month, 13, 54
	EMWriteScreen "05", 13, 57
	EMWriteScreen bene_year, 13, 60
	EMWriteScreen "________", 13, 68
	EMWriteScreen inc_amount, 13, 68
	
	IF pay_freq = "2" OR pay_freq = "3" THEN
		EMWriteScreen retro_month, 14, 25
		EMWriteScreen "19", 14, 28
		EMWriteScreen retro_year, 14, 31
		EMWriteScreen "________", 14, 39
		EMWriteScreen inc_amount, 14, 39
		EMWriteScreen bene_month, 14, 54
		EMWriteScreen "19", 14, 57
		EMWriteScreen bene_year, 14, 60
		EMWriteScreen "________", 14, 68
		EMWriteScreen inc_amount, 14, 68
	ELSEIF pay_freq = "4" THEN
		EMWriteScreen retro_month, 14, 25
		EMWriteScreen "12", 14, 28
		EMWriteScreen retro_year, 14, 31
		EMWriteScreen "________", 14, 39
		EMWriteScreen inc_amount, 14, 39
		EMWriteScreen retro_month, 15, 25
		EMWriteScreen "19", 15, 28
		EMWriteScreen retro_year, 15, 31
		EMWriteScreen "________", 15, 39
		EMWriteScreen inc_amount, 15, 39
		EMWriteScreen retro_month, 16, 25
		EMWriteScreen "26", 16, 28
		EMWriteScreen retro_year, 16, 31
		EMWriteScreen "________", 16, 39
		EMWriteScreen inc_amount, 16, 39
		EMWriteScreen bene_month, 14, 54
		EMWriteScreen "12", 14, 57
		EMWriteScreen bene_year, 14, 60
		EMWriteScreen "________", 14, 68
		EMWriteScreen inc_amount, 14, 68
		EMWriteScreen bene_month, 15, 54 
		EMWriteScreen "19", 15, 57 
		EMWriteScreen bene_year, 15, 60 
		EMWriteScreen "________", 15, 68 
		EMWriteScreen inc_amount, 15, 68 
		EMWriteScreen bene_month, 16, 54
		EMWriteScreen "26", 16, 57
		EMWriteScreen bene_year, 16, 60
		EMWriteScreen "________", 16, 68
		EMWriteScreen inc_amount, 16, 68
	END IF

	'=====determines if the benefit month is current month + 1 and dumps information into the HC income estimator
	IF (bene_month * 1) = (datepart("M", date) + 1) THEN
		EMWriteScreen "X", 6, 56
		EMWriteScreen "________", 9, 65
		EMWriteScreen inc_amount, 9, 65
		EMWriteScreen pay_freq, 10, 63
		transmit
		transmit
	END IF

END FUNCTION

FUNCTION write_panel_to_maxis_DIET(mfip1, mfipv1, mfip2, mfipv2, msa1, msav1, msa2, msav2, msa3, msav3, msa4, msav4)
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

FUNCTION write_panel_to_maxis_MMSA(liv_arr, cont_elig, spous_inc, shared_hous)
	EMWriteScreen liv_arr, 7, 54
	EMWriteScreen cont_elig, 9, 54
	EMWriteScreen spous_inc, 12, 62
	EMWriteScreen shared_hous, 14, 62
	transmit
END FUNCTION

FUNCTION write_panel_maxis_ADDR(addr1, addr2, city, zip, res_co, verif1, homeless, ind_res, res_name, mail1, mail2, mailcity, mailzip, ph1, ph2, ph3)
	call create_maxis_friendly_date(appl_date, 0, 4, 43)
	EMWriteScreen addr1, 6, 43
	EMWriteScreen addr2, 7, 43
	EMWriteScreen city, 8, 43
	EMWriteScreen "MN", 8, 66
	EMWriteScreen zip, 9, 43
	EMWriteScreen res_co, 9, 66
	EMWriteScreen verif1, 9, 74
	EMWriteScreen homeless, 10, 43
	EMWriteScreen ind_res, 10, 74
	EMWriteScreen res_name, 11, 74
	EMWriteScreen mail1, 13, 43
	EMWriteScreen mail2, 14, 43
	EMWriteScreen mailcity, 15, 43
	EMWriteScreen "MN", 16, 43
	EMWriteScreen mailzip, 16, 52
	EMWriteScreen left(ph1, 3), 17, 45
	EMWriteScreen right(left(ph1, 6), 3), 17, 51
	EMWriteScreen right(ph1, 4), 17, 55
	EMWriteScreen left(ph2, 3), 18, 45
	EMWriteScreen right(left(ph2, 6), 3), 18, 51
	EMWriteScreen right(ph2, 4), 18, 51
END FUNCTION
