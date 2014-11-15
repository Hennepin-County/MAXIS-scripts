
'PARE--------------------------------------------------------------------------------------------------------------------------------------------------------
PARE_child_1 = "03"
PARE_child_2 = "04"
PARE_child_3 = ""
PARE_child_4 = ""
PARE_child_5 = ""
PARE_child_6 = ""
PARE_child_1_relation = "1"
PARE_child_2_relation = "1"
PARE_child_3_relation = ""
PARE_child_4_relation = ""
PARE_child_5_relation = ""
PARE_child_6_relation = ""
PARE_child_1_verif = "OT"
PARE_child_2_verif = "BC"
PARE_child_3_verif = ""
PARE_child_4_verif = ""
PARE_child_5_verif = ""
PARE_child_6_verif = ""

FUNCTION write_panel_to_maxis_PARE(PARE_child_1, PARE_child_1_relation, PARE_child_1_verif, PARE_child_2, PARE_child_2_relation, PARE_child_2_verif, PARE_child_3, PARE_child_3_relation, PARE_child_3_verif, PARE_child_4, PARE_child_4_relation, PARE_child_4_verif, PARE_child_5, PARE_child_5_relation, PARE_child_5_verif, PARE_child_6, PARE_child_6_relation, PARE_child_6_verif)
	Call navigate_to_screen("STAT", "PARE") 
	call create_if_nonexistant
	call create_if_nonexistant
		EMWritescreen PARE_child_1, 8, 24
		EMWritescreen PARE_child_1_relation, 8, 53
		EMWritescreen PARE_child_1_verif, 8, 71
		EMWritescreen PARE_child_2, 9, 24
		EMWritescreen PARE_child_2_relation, 9, 53
		EMWritescreen PARE_child_2_verif, 9, 71
		EMWritescreen PARE_child_3, 10, 24
		EMWritescreen PARE_child_3_relation, 10, 53
		EMWritescreen PARE_child_3_verif, 10, 71
		EMWritescreen PARE_child_4, 11, 24
		EMWritescreen PARE_child_4_relation, 11, 53
		EMWritescreen PARE_child_4_verif, 11, 71
		EMWritescreen PARE_child_5, 12, 24
		EMWritescreen PARE_child_5_relation, 12, 53
		EMWritescreen PARE_child_5_verif, 12, 71
		EMWritescreen PARE_child_6, 13, 24
		EMWritescreen PARE_child_6_relation, 13, 53
		EMWritescreen PARE_child_6_verif, 13, 71
	transmit
end function

'ACUT-------------------------------------------------------------------------------------------------------------------------------------------------------------
ACUT_shared = "N"
ACUT_heat = "137"
ACUT_air = ""
ACUT_electric = "145"
ACUT_fuel = "135"
ACUT_garbage = ""
ACUT_water = ""
ACUT_sewer = ""
ACUT_other = ""
ACUT_phone = "yes" 
ACUT_heat_verif = "y" 
ACUT_air_verif = "y"
ACUT_electric_verif = ""
ACUT_fuel_verif = ""
ACUT_garbage_verif = ""
ACUT_water_verif = ""
ACUT_sewer_verif = ""
ACUT_other_verif = ""

'THE FUNCTION
FUNCTION write_panel_to_maxis_ACUT(ACUT_shared, ACUT_heat, ACUT_air, ACUT_electric, ACUT_fuel, ACUT_garbage, ACUT_water, ACUT_sewer, ACUT_other, ACUT_phone, ACUT_heat_verif, ACUT_air_verif, ACUT_electric_verif, ACUT_fuel_verif, ACUT_garbage_verif, ACUT_water_verif, ACUT_sewer_verif, ACUT_other_verif)

call navigate_to_screen("STAT", "ACUT")
call create_if_nonexistant
 EMWritescreen ACUT_shared, 6, 42
 EMWritescreen ACUT_heat, 10, 61
 EMWritescreen ACUT_air, 11, 61
 EMWritescreen ACUT_electric, 12, 61
 EMWritescreen ACUT_fuel, 13, 61
 EMWritescreen ACUT_garbage, 14, 61
 EMWritescreen ACUT_water, 15, 61
 EMWritescreen ACUT_sewer, 16, 61
 EMWritescreen ACUT_other, 17, 61
 EMWritescreen ACUT_heat_verif, 10, 55
 EMWritescreen ACUT_air_verif, 11, 55
 EMWritescreen ACUT_electric_verif, 12, 55
 EMWritescreen ACUT_fuel_verif, 13, 55
 EMWritescreen ACUT_garbage_verif, 14, 55
 EMWritescreen ACUT_water_verif, 15, 55
 EMWritescreen ACUT_sewer_verif, 16, 55
 EMWritescreen ACUT_other_verif, 17, 55
 EMWritescreen Left(ACUT_phone, 1), 18, 55
transmit
end function

'PREG----------------------------------------------------------------------------------------------------------------------------------------------------------------
PREG_conception_date_ver = "y"
PREG_conception_date = "10/19/2014"
PREG_third_trimester_ver = "?"
PREG_due_date = "08/05/14"
PREG_multiple_birth = ""



FUNCTION write_panel_to_maxis_PREG(PREG_conception_date, PREG_conception_date_ver, PREG_third_trimester_ver,PREG_due_date, PREG_multiple_birth)

 call navigate_to_screen("STAT", "PREG")
 call create_if_nonexistant
	EMWritescreen "NN", 20, 79
	transmit
	call create_MAXIS_friendly_date(PREG_conception_date, 1, 6, 53)
	call create_MAXIS_friendly_date(PREG_due_date, 1, 10, 53)
	EMWritescreen PREG_conception_date_ver, 6, 75
	EMWritescreen PREG_third_trimester_ver, 8, 75
	EMWritescreen PREG_multiple_birth, 14, 53
 transmit
end function

'call write_panel_to_maxis_PREG(PREG_conception_date, PREG_conception_date_ver, PREG_third_trimester_ver,PREG_due_date, PREG_multiple_birth)

'HEST--------------------------------------------------------------------------------------------------------------------------------------------------
HEST_FS_choice_date = "11/01/14"
HEST_first_month = "450" 'actual expense in initial month
HEST_heat_air_retro = "yes"
HEST_electric_retro = ""
HEST_phone_retro = ""
HEST_heat_air_pro = "Y"
HEST_electric_pro = ""
HEST_phone_pro = ""



FUNCTION write_panel_to_maxis_HEST(HEST_FS_choice_date, HEST_first_month, HEST_heat_air_retro, HEST_electric_retro, HEST_phone_retro, HEST_heat_air_pro, HEST_electric_pro, HEST_phone_pro)
 call navigate_to_screen("STAT", "HEST")
 call create_if_nonexistant
	call create_MAXIS_friendly_date(HEST_FS_choice_date, 1, 07, 40)
	EMWritescreen HEST_first_month, 8, 61 
	'Filling in the #/FS units field (always 01)
     If ucase(left(HEST_heat_air_retro, 1)) = "Y" then EMWritescreen "01", 13, 42
	 If ucase(left(HEST_heat_air_pro, 1)) = "Y" then EMWritescreen "01", 13, 68
	 If ucase(left(HEST_electric_retro, 1)) = "Y" then EMWritescreen "01", 14, 42
	 If ucase(left(HEST_electric_pro, 1)) = "Y" then EMWritescreen "01", 14, 68
	 If ucase(left(HEST_phone_retro, 1)) = "Y" then EMWritescreen "01", 15, 42
	 If ucase(left(HEST_phone_pro, 1)) = "Y" then EMWritescreen "01", 15, 68
	EMWritescreen left(HEST_heat_air_retro, 1), 13, 34
	EMWritescreen left(HEST_electric_retro, 1), 14, 34
	EMWritescreen left(HEST_phone_retro, 1), 15, 34
	EMWritescreen left(HEST_heat_air_pro, 1), 13, 60
	EMWritescreen left(HEST_electric_pro, 1), 14, 60
	EMWritescreen left(HEST_phone_pro, 1), 15, 60
transmit
end function

'call write_panel_to_maxis_HEST(HEST_FS_choice_date, HEST_first_month, HEST_heat_air_retro, HEST_electric_retro, HEST_phone_retro, HEST_heat_air_pro, HEST_electric_pro, HEST_phone_pro)

'SIBL-------------------------------------------------------------------------------------------------------------------------------------------------------------
SIBL_group_1 = "2, 11, 13"
SIBL_group_2 = ""
SIBL_group_3 = ""

FUNCTION write_panel_to_maxis_SIBL(SIBL_group_1, SIBL_group_2, SIBL_group_3)
 call navigate_to_screen("STAT", "SIBL")
 call create_if_nonexistant
	If SIBL_group_1 <> "" then 
		EMWritescreen "01", 13, 28
		SIBL_group_1 = relace(SIBL_group_1, " ", "") 'Removing spaces
		SIBL_group_1 = split(SIBL_group_1, ",") 'Splits the sibling group value into an array by commas
		SIBL_col = 39
		For Each SIBL_group_member in SIBL_group_1 'Writes the member numbers onto the group line
			EMWritescreen SIBL_group_member, 13, SIBL_col
			SIBL_col = SIBL_col + 5
		Next
	
	If SIBL_group_2 <> "" then
		EMWritescreen "02", 14, 28
		SIBL_group_2 = relace(SIBL_group_2, " ", "")
		SIBL_group_2 = split(SIBL_group_2, ",")
		SIBL_col = 39
		For Each SIBL_group_member in SIBL_group_2
			EMWritescreen SIBL_group_member, 14, SIBL_col
			SIBL_col = SIBL_col + 5
		Next
	
	If SIBL_group_3 <> "" then
		EMWritescreen "03", 15, 28
		SIBL_group_2 = relace(SIBL_group_3, " ", "")
		SIBL_group_2 = split(SIBL_group_3, ",")
		SIBL_col = 39
		For Each SIBL_group_member in SIBL_group_3
			EMWritescreen SIBL_group_member, 14, SIBL_col
			SIBL_col = SIBL_col + 5
		Next
			
	transmit
end function

'EMPS--------------------------------------------------------------------------------------------------------------------------------------------------------------
EMPS_orientation_date = ""
EMPS_orientation_attended = "y" 'this is a Y/N field
EMPS_good_cause = ""
EMPS_sanc_begin = ""
EMPS_sanc_end = ""
'The following 5 variables must be entered or the panel will error!
EMPS_memb_at_home = "n" 'member required at home for special medical criteria
EMPS_care_family = "n" 'member required at home for care of ill family member
EMPS_crisis = "n" 'member experiencing personal/family crisis
EMPS_hard_employ = "no" 'member meets hard to employ category
EMPS_under1 = "n" 'FT care of child under 1

EMPS_DWP_date = "" 'DWP plan date

FUNCTION write_panel_to_maxis_EMPS(EMPS_orientation_date, EMPS_orientation_attended, EMPS_good_cause, EMPS_sanc_begin, EMPS_sanc_end, EMPS_memb_at_home, EMPS_care_family, EMPS_crisis, EMPS_hard_employ, EMPS_under1, EMPS_DWP_date)
call navigate_to_screen("STAT", "EMPS")
call create_if_nonexistant
	If EMPS_orientation_date <> "" then call create_maxis_friendly_date(EMPS_orientation_date, 0, 5, 39) 'enter orientation date
	EMWritescreen left(EMPS_orientation_attended, 1), 5, 65 
	EMWritescreen EMPS_good_cause, 5, 79
	If EMPS_sanc_begin <> "" then call create_maxis_friendly_date(EMPS_sanc_begin, 1, 6, 39) 'Sanction begin date
	If EMPS_sanc_end <> "" then call create_maxis_friendly_date(EMPS_sanc_end, 1, 6, 65) 'Sanction end date
	EMWritescreen left(EMPS_memb_at_home, 1), 8, 76
	EMWritescreen left(EMPS_care_family, 1), 9, 76
	EMWritescreen left(EMPS_crisis, 1), 10, 76
	EMWritescreen EMPS_hard_employ, 11, 76
	EMWritescreen left(EMPS_under1, 1), 12, 76
	EMWritescreen "n", 13, 76 'enters n for child under 12 weeks
	If EMPS_DWP_date <> "" then call create_maxis_friendly_date(EMPS_DWP_date, 1, 17, 40) 'DWP plan date

End Function	

'call write_panel_to_maxis_EMPS(EMPS_orientation_date, EMPS_orientation_attended, EMPS_good_cause, EMPS_sanc_begin, EMPS_sanc_end, EMPS_memb_at_home, EMPS_care_family, EMPS_crisis, EMPS_hard_employ, EMPS_under1, EMPS_DWP_date)

'DCEX------------------------------------------------------------------------------------------------------------------------------------------------


DCEX_provider = "" 
DCEX_reason = ""
DCEX_subsidy = ""
DCEX_child_number1 = ""
DCEX_child_number2 = ""
DCEX_child_number3 = ""
DCEX_child_number4 = ""
DCEX_child_number5 = ""
DCEX_child_number6 = ""
DCEX_child_number1_ver = ""
DCEX_child_number2_ver = ""
DCEX_child_number3_ver = ""
DCEX_child_number4_ver = ""
DCEX_child_number5_ver = ""
DCEX_child_number6_ver = ""
DCEX_child_number1_retro = ""
DCEX_child_number2_retro = ""
DCEX_child_number3_retro = ""
DCEX_child_number4_retro = ""
DCEX_child_number5_retro = ""
DCEX_child_number6_retro = ""
DCEX_child_number1_pro = ""
DCEX_child_number2_pro = ""
DCEX_child_number3_pro = ""
DCEX_child_number4_pro = ""
DCEX_child_number5_pro = ""
DCEX_child_number6_pro = ""
              
FUNCTION write_panel_to_maxis_DCEX(DCEX_provider, DCEX_reason, DCEX_subsidy, DCEX_child_number1, DCEX_child_number1_ver, DCEX_child_number1_retro, DCEX_child_number1_pro, DCEX_child_number2, DCEX_child_number2_ver, DCEX_child_number2_retro, DCEX_child_number2_pro, DCEX_child_number3, DCEX_child_number3_ver, DCEX_child_number3_retro, DCEX_child_number3_pro, DCEX_child_number4, DCEX_child_number4_ver, DCEX_child_number4_retro, DCEX_child_number4_pro, DCEX_child_number5, DCEX_child_number5_ver, DCEX_child_number5_retro, DCEX_child_number5_pro, DCEX_child_number6, DCEX_child_number6_ver, DCEX_child_number6_retro, DCEX_child_number6_pro)
call navigate_to_screen("STAT", "DCEX") 
call create_if_nonexistant
	EMWritescreen DCEX_provider, 6, 47
	EMWritescreen DCEX_reason, 7, 44
	EMWritescreen DCEX_subsidy, 8, 44
	EMWritescreen DCEX_child_number1, 11, 29
	EMWritescreen DCEX_child_number2, 12, 29
	EMWritescreen DCEX_child_number3, 13, 29
	EMWritescreen DCEX_child_number4, 14, 29
	EMWritescreen DCEX_child_number5, 15, 29
	EMWritescreen DCEX_child_number6, 16, 29
	EMWritescreen DCEX_child_number1_ver, 11, 41
	EMWritescreen DCEX_child_number2_ver, 12, 41
	EMWritescreen DCEX_child_number3_ver, 13, 41
	EMWritescreen DCEX_child_number4_ver, 14, 41
	EMWritescreen DCEX_child_number5_ver, 15, 41
	EMWritescreen DCEX_child_number6_ver, 16, 41
	EMWritescreen DCEX_child_number1_retro, 11, 48
	EMWritescreen DCEX_child_number2_retro, 12, 48
	EMWritescreen DCEX_child_number3_retro, 13, 48
	EMWritescreen DCEX_child_number4_retro, 14, 48
	EMWritescreen DCEX_child_number5_retro, 15, 48
	EMWritescreen DCEX_child_number6_retro, 16, 48
	EMWritescreen DCEX_child_number1_pro, 11, 63
	EMWritescreen DCEX_child_number2_pro, 11, 63
	EMWritescreen DCEX_child_number3_pro, 11, 63
	EMWritescreen DCEX_child_number4_pro, 11, 63
	EMWritescreen DCEX_child_number5_pro, 11, 63
	EMWritescreen DCEX_child_number6_pro, 11, 63
transmit

End Function	

'call write_panel_to_maxis_DCEX(DCEX_provider, DCEX_reason, DCEX_subsidy, DCEX_child_number1, DCEX_child_number1_ver, DCEX_child_number1_retro, DCEX_child_number1_pro, DCEX_child_number2, DCEX_child_number2_ver, DCEX_child_number2_retro, DCEX_child_number2_pro, DCEX_child_number3, DCEX_child_number3_ver, DCEX_child_number3_retro, DCEX_child_number3_pro, DCEX_child_number4, DCEX_child_number4_ver, DCEX_child_number4_retro, DCEX_child_number4_pro, DCEX_child_number5, DCEX_child_number5_ver, DCEX_child_number5_retro, DCEX_child_number5_pro, DCEX_child_number6, DCEX_child_number6_ver, DCEX_child_number6_retro, DCEX_child_number6_pro)

'SHEL---------------------------------------------------------------------------------------------------------------------------------------------
SHEL_subsidized = ""
SHEL_shared = ""
SHEL_paid_to = "Yossarian"
SHEL_rent_retro = ""
SHEL_rent_retro_ver = ""
SHEL_rent_pro = ""
SHEL_rent_pro_ver
SHEL_lot_rent_retro = ""
SHEL_lot_rent_retro_ver = ""
SHEL_lot_rent_pro = ""
SHEL_lot_rent_pro_ver = ""
SHEL_mortgage_retro = ""
SHEL_mortgage_retro_ver = ""
SHEL_mortgage_pro = ""
SHEL_mortgage_pro_ver = ""
SHEL_insur_retro = ""
SHEL_insur_retro_ver = ""
SHEL_insur_pro = ""
SHEL_insur_pro_ver = ""
SHEL_taxes_retro = ""
SHEL_taxes_retro_ver = ""
SHEL_taxes_pro = ""
SHEL_taxes_pro_ver = ""
SHEL_room_retro = ""
SHEL_room_retro_ver = ""
SHEL_room_pro = ""
SHEL_room_pro_ver = ""
SHEL_garage_retro = ""
SHEL_garage_retro_ver = ""
SHEL_garage_pro = ""
SHEL_garage_pro_ver = ""
SHEL_subsidy_retro = ""
SHEL_subsidy_retro_ver = ""
SHEL_subsidy_pro = ""
SHEL_subsidy_pro_ver = ""

FUNCTION write_panel_to_maxis_SHEL(SHEL_subsidized, SHEL_shared, SHEL_paid_to, SHEL_rent_retro, SHEL_rent_retro_ver, SHEL_rent_pro, SHEL_rent_pro_ver, SHEL_lot_rent_retro, SHEL_lot_rent_retro_ver, SHEL_lot_rent_pro, SHEL_lot_rent_pro_ver, SHEL_mortgage_retro, SHEL_mortgage_retro_ver, SHEL_mortgage_pro, SHEL_mortgage_pro_ver, SHEL_insur_retro, SHEL_insur_retro_ver, SHEL_insur_pro, SHEL_insur_pro_ver, SHEL_taxes_retro, SHEL_taxes_retro_ver, SHEL_taxes_pro SHEL_taxes_pro_ver, SHEL_room_retro, SHEL_room_retro_ver, SHEL_room_pro, SHEL_room_pro_ver, SHEL_garage_retro, SHEL_garage_retro_ver, SHEL_garage_pro, SHEL_garage_pro_ver, SHEL_subsidy_retro, SHEL_subsidy_retro_ver, SHEL_subsidy_pro, SHEL_subsidy_pro_ver)
call navigate_to_screen("STAT", "SHEL")
call create_if_nonexistant
	EMWritescreen SHEL_subsidized, 6, 42
	EMWritescreen SHEL_shared, 6, 60
	EMWritescreen SHEL_paid_to, 7, 46
	EMWritescreen SHEL_rent_retro, 11, 37
	EMWritescreen SHEL_rent_retro_ver, 11, 48
	EMWritescreen SHEL_rent_pro, 11, 56
	EMWritescreen SHEL_rent_pro_ver, 11, 67
	EMWritescreen SHEL_lot_rent_retro, 12, 37
	EMWritescreen SHEL_lot_rent_retro_ver, 12, 48
	EMWritescreen SHEL_lot_rent_pro, 12, 56
	EMWritescreen SHEL_lot_rent_pro_ver, 12, 67
	EMWritescreen SHEL_mortgage_retro, 13, 37
	EMWritescreen SHEL_mortgage_retro_ver, 13, 48
	EMWritescreen SHEL_mortgage_pro, 13, 56
	EMWritescreen SHEL_insur_retro, 14, 37 
	EMWritescreen SHEL_insur_retro_ver, 14, 48
	EMWritescreen SHEL_insur_pro, 14, 56
	EMWritescreen SHEL_insur_pro_ver, 14, 67
	EMWritescreen SHEL_taxes_retro, 15, 37
	EMWritescreen SHEL_taxes_retro_ver, 15, 48
	EMWritescreen SHEL_taxes_pro, 15, 56
	EMWritescreen SHEL_taxes_pro_ver, 15, 67
	EMWritescreen SHEL_room_retro, 16, 37
	EMWritescreen SHEL_room_retro_ver, 16, 48
	EMWritescreen SHEL_room_pro, 16, 56
	EMWritescreen SHEL_room_pro_ver, 16, 67
	EMWritescreen SHEL_garage_retro, 17, 37
	EMWritescreen SHEL_garage_retro_ver, 17, 48
	EMWritescreen SHEL_garage_pro, 17, 56
	EMWritescreen SHEL_garage_pro_ver, 17, 67
	EMWritescreen SHEL_subsidy_retro, 18, 37
	EMWritescreen SHEL_subsidy_retro_ver, 18, 48
	EMWritescreen SHEL_subsidy_pro, 18, 56
	EMWritescreen SHEL_subsidy_pro, 18, 67

transmit

End Function	
