'writing the write functions

'->look at asset ratios! default to "1 1" even on the sheet?

Function write_panel_to_maxis_CARS(cars_type, cars_year, cars_make, cars_model, cars_trade_in, cars_loan, cars_value_source, cars_ownership_ver, cars_amount_owed, cars_amount_owed_ver, cars_date, cars_owed_as_of, cars_use, cars_HC_benefit, cars_joint_owner, cars_share_ratio)
	Call Navigate_to_screen(STAT, CARS)
	Emwritescreen cars_type, 6, 43
	Emwritescreen cars_year, 8, 31
	Emwritescreen cars_make, 8, 43
	Emwritescreen cars_model, 8, 66
	Emwritescreen cars_trade_in, 9, 45
	Emwritescreen cars_loan, 9, 62
	Emwritescreen cars_value_source, 9, 80
	Emwritescreen cars_ownership_ver, 10, 60
	Emwritescreen cars_amount_owed, 12, 45
	Emwritescreen cars_amount_owed_ver, 12, 60
	call create_maxis_friendly_date(cars_date, 0, 13, 43)
	Emwritescreen cars_owed_as_of, 13, 43
	Emwritescreen cars_use, 15, 43
	Emwritescreen cars_HC_benefit, 15, 76
	Emwritescreen cars_joint_owner
	Emwritescreen cars_share_ratio
End Function

Function write_panel_to_maxis_ACCT(acct_type, acct_numb, acct_location, acct_balance, acct_bal_ver, acct_date, acct_withdraw, acct_cash_count, acct_snap_count, acct_HC_count, acct_GRH_count, acct_IV_count, acct_joint_owner, acct_share_ratio, acct_interest_date_mo, acct_interest_date_yr)
	Call Navigate_to_screen(STAT, ACCT)
	Emwritescreen acct_type, 6, 44
	Emwritescreen acct_numb, 7, 44
	Emwritescreen acct_location, 8, 44
	Emwritescreen acct_balance, 10, 46
	Emwritescreen acct_bal_ver, 10, 63
	call create_maxis_friendly_date(acct_date, 0, 11, 44)
	Emwritescreen acct_withdraw, 12, 46
	Emwritescreen acct_cash_count, 14, 50
	Emwritescreen acct_snap_count, 14, 57
	Emwritescreen acct_HC_count, 14, 64
	Emwritescreen acct_GRH_count, 14, 72
	Emwritescreen acct_IV_count, 14, 80
	Emwritescreen acct_joint_owner, 15, 44
	Emwritescreen acct_share_ratio, 15, 76
	Emwritescreen acct_interest_date_mo, 17, 57
	Emwritescreen acct_interest_date_yr, 17, 60
End Function

Function write_panel_to_maxis_CASH(cash_amount)
	Call navigate_to_screen(STAT, CASH)
	Emwritescreen cash_amount, 8, 39
End Function

Function write_panel_to_maxis_OTHR(othr_type, othr_cash_value, othr_cash_value_ver, othr_owed, othr_owed_ver, othr_date, othr_cash_count, othr_SNAP_count, othr_HC_count, othr_IV_count, othr_joint, othr_share_ratio)
	Call navigate_to_screen(STAT, OTHR)
	Emwritescreen othr_type, 6, 40
	Emwritescreen othr_cash_value, 8, 40
	Emwritescreen othr_cash_value_ver, 8, 57
	Emwritescreen othr_owed, 9, 40
	Emwritescreen othr_owed_ver, 9, 57
	call create_maxis_friendly_date(othr_date, 0, 10, 39)
	Emwritescreen othr_cash_count, 12, 50
	Emwritescreen othr_SNAP_count, 12, 57
	Emwritescreen othr_HC_count, 12, 64
	Emwritescreen othr_IV_count, 12, 73
	Emwritescreen othr_joint, 13, 44
	Emwritescreen othr_share_ratio, 15, 50
End Function

Function write_panel_to_maxis_SECU(secu_type, secu_pol_numb, secu_name, secu_cash_val, secu_date, secu_cash_ver, secu_face_val, secu_withdraw, secu_cash_count, secu_SNAP_count, secu_HC_count, secu_GRH_count, secu_IV_count, secu_joint, secu_share_ratio)
	Call navigate_to_screen(STAT, SECU)
	Emwritescreen secu_type, 6, 50
	Emwritescreen secu_pol_numb, 7, 50
	Emwritescreen secu_name, 8, 50
	Emwritescreen secu_cash_val, 10, 52
	call create_maxis_friendly_date(secu_date, 0, 11, 35)
	Emwritescreen secu_cash_ver, 11, 50
	Emwritescreen secu_face_val, 12, 52
	Emwritescreen secu_withdraw, 13, 52
	Emwritescreen secu_cash_count, 15, 50
	Emwritescreen secu_SNAP_count, 15, 57
	Emwritescreen secu_HC_count, 15, 64
	Emwritescreen secu_GRH_count, 15, 72
	Emwritescreen secu_IV_count, 15, 80
	Emwritescreen secu_joint, 16, 44
	Emwritescreen secu_share_ratio, 16, 76
End Function

Function write_panel_to_maxis_REST(rest_type, rest_type_ver, rest_market, rest_market_ver, rest_owed, rest_owed_ver, rest_date, rest_status, rest_joint, rest_share_ratio, rest_agreement_date)
	Call navigate_to_screen(STAT, REST)
	Emwritescreen rest_type, 6, 39
	Emwritescreen rest_type_ver, 6, 62
	Emwritescreen rest_market, 8, 41
	Emwritescreen rest_market_ver, 8, 62
	Emwritescreen rest_owed, 9, 41
	Emwritescreen rest_owed_ver, 9, 62
	call create_maxis_friendly_date(rest_date, 0, 10, 39)
	Emwritescreen rest_status, 12, 54
	Emwritescreen rest_joint, 13, 54
	Emwritescreen rest_share_ratio, 14, 54
	call create_maxis_friendly_date(rest_agreement_date, 0, 16, 62)
End Function

Function write_panel_to_maxis_DISA(disa_begin_date, disa_end_date, disa_cert_begin, disa_cert_end, disa_wavr_begin, disa_wavr_end, disa_grh_begin, disa_grh_end, disa_cash_status, disa_cash_status_ver, disa_snap_status, disa_snap_status_ver, disa_hc_status, disa_hc_status_ver, disa_waiver, disa_drug_alcohol)
	Call navigate_to_screen(STAT, DISA)
	call create_maxis_friendly_date(disa_begin_date, 0, 6, 47)
	call create_maxis_friendly_date(disa_end_date, 0, 6, 69)
	call create_maxis_friendly_date(disa_cert_begin, 0, 7, 47)
	call create_maxis_friendly_date(disa_cert_end, 0, 7, 69)
	call create_maxis_friendly_date(disa_wavr_begin, 0, 8, 47)
	call create_maxis_friendly_date(disa_wavr_end, 0, 8, 69)
	call create_maxis_friendly_date(disa_ghr_begin, 0, 9, 47)
	call create_maxis_friendly_date(disa_ghr_end, 0, 9, 69)
	Emwritescreen disa_cash_status, 11, 59
	Emwritescreen disa_cash_status_ver, 11, 69
	Emwritescreen disa_snap_status, 12, 59
	Emwritescreen disa_snap_status_ver, 12, 69
	Emwritescreen disa_hc_status, 13, 59
	Emwritescreen disa_hc_status_ver, 13, 69
	Emwritescreen disa_waiver, 14, 59
	Emwritescreen disa_1619, 16, 59
	Emwritescreen disa_drug_alcohol, 18, 69
End Function
	
Function write_panel_to_maxis_PBEN()
	Call navigate_to_screen(STAT, PBEN)
	Emreadscreen pben_row_check, 2, 8, 24
	If pben_row_check = "  " THEN
		Emwritescreen pben_type, 8, 24
		call create_maxis_friendly_date(pben_referal_date, 0, 8, 40)
		call create_maxis_friendly_date(pben_appl_date, 0, 8, 51)
		Emwritescreen pben_appl_ver, 8, 62
		call create_maxis_friendly_date(pben_IAA_date, 0, 8, 66)
		Emwritescreen pben_disp, 8, 77
	else 
		EMreadscreen pben_row_check, 2, 9, 24
		IF pben_row_check = "  " THEN
		'second pben row
			Emwritescreen pben_type, 9, 24
			call create_maxis_friendly_date(pben_referal_date, 0, 9, 40)
			call create_maxis_friendly_date(pben_appl_date, 0, 9, 51)
			Emwritescreen pben_appl_ver, 9, 62
			call create_maxis_friendly_date(pben_IAA_date, 0, 9, 66)
			Emwritescreen pben_disp, 9, 77
		else
		Emreadscreen pben_row_check, 2, 10, 24
			IF pben-row_check = "  " THEN
			'third pben row
				Emwritescreen pben_type, 10, 24
				call create_maxis_friendly_date(pben_referal_date, 0, 10, 40)
				call create_maxis_friendly_date(pben_appl_date, 0, 10, 51)
				Emwritescreen pben_appl_ver, 10, 62
				call create_maxis_friendly_date(pben_IAA_date, 0, 10, 66)
				Emwritescreen pben_disp, 10, 77
			END IF
		END IF
	END IF
End Function
