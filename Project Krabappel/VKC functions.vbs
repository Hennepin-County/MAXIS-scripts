'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-MAXIS-Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'VARIABLES
case_number = "210668"
reference_number = "01"
EMConnect ""

'--------------Going to MEMB, this should be in the MEMB section of the script
call navigate_to_screen("STAT", "MEMB")

'Enters reference_number into MEMB and transmit
EMWriteScreen reference_number, 20, 76
transmit

'Grabs SSN for later use
EMReadScreen SSN_first, 3, 7, 42
EMReadScreen SSN_mid, 2, 7, 46
EMReadScreen SSN_last, 4, 7, 49

'----------END MEMB PIECES

'IMIG----------------------------------------------------------------------------------------------------------------------------------
IMIG_imigration_status = "21"
IMIG_entry_date = "11/13/2013"
IMIG_status_date = "11/13/2013"
IMIG_status_ver = "OT"
IMIG_status_LPR_adj_from = ""
IMIG_nationality = "MX"

Function write_panel_to_MAXIS_IMIG(IMIG_imigration_status, IMIG_entry_date, IMIG_status_date, IMIG_status_ver, IMIG_status_LPR_adj_from, IMIG_nationality)
	call navigate_to_screen("STAT", "IMIG")
	call ERRR_screen_check
	call create_panel_if_nonexistent
	call create_MAXIS_friendly_date(date, 0, 5, 45)						'Writes actual date, needs to add 2000 as this is weirdly a 4 digit year
	EMWriteScreen datepart("yyyy", date), 5, 51
	EMWriteScreen IMIG_imigration_status, 6, 45							'Writes imig status
	call create_MAXIS_friendly_date(IMIG_entry_date, 0, 7, 45)			'Enters year as a 2 digit number, so have to modify manually
	EMWriteScreen datepart("yyyy", IMIG_entry_date), 7, 51
	call create_MAXIS_friendly_date(IMIG_status_date, 0, 7, 71)			'Enters year as a 2 digit number, so have to modify manually
	EMWriteScreen datepart("yyyy", IMIG_status_date), 7, 77
	EMWriteScreen IMIG_status_ver, 8, 45								'Enters status ver
	EMWriteScreen IMIG_status_LPR_adj_from, 9, 45						'Enters status LPR adj from
	EMWriteScreen IMIG_nationality, 10, 45								'Enters nationality
	transmit
	transmit
End function

'call write_panel_to_MAXIS_IMIG(imigration_status, entry_date, status_date, status_ver, status_LPR_adj_from, nationality)

'SPON----------------------------------------------------------------------------------------------------------------------------------
SPON_type = "IL"
SPON_ver = "Y"
SPON_name = "Bob Loblaw"
SPON_state = "MN"

Function write_panel_to_MAXIS_SPON(SPON_type, SPON_ver, SPON_name, SPON_state)
	call navigate_to_screen("STAT", "SPON")
	call ERRR_screen_check
	call create_panel_if_nonexistent
	EMWriteScreen SPON_type, 6, 38
	EMWriteScreen SPON_ver, 6, 62
	EMWriteScreen SPON_name, 8, 38
	EMWriteScreen SPON_state, 10, 62
	transmit
End function

'call write_panel_to_MAXIS_SPON(SPON_type, SPON_ver, SPON_name, SPON_state)

'DSTT----------------------------------------------------------------------------------------------------------------------------------
DSTT_ongoing_income = "Y"
DSTT_HH_income_stop_date = "11/13/2014"
DSTT_income_expected_amt = "300"

Function write_panel_to_MAXIS_DSTT(DSTT_ongoing_income, DSTT_HH_income_stop_date, DSTT_income_expected_amt)
	call navigate_to_screen("STAT", "DSTT")
	call ERRR_screen_check
	call create_panel_if_nonexistent
	EMWriteScreen DSTT_ongoing_income, 6, 69
	call create_MAXIS_friendly_date(HH_income_stop_date, 0, 9, 69)
	EMWriteScreen income_expected_amt, 12, 71
End function

'call write_panel_to_MAXIS_DSTT(DSTT_ongoing_income, HH_income_stop_date, income_expected_amt)

'EMMA----------------------------------------------------------------------------------------------------------------------------------

EMMA_medical_emergency = "02"
EMMA_health_consequence = "01"
EMMA_verification = "OT"
EMMA_begin_date = "10/14/2013"
EMMA_end_date = "11/14/2013"

Function write_panel_to_MAXIS_EMMA(EMMA_medical_emergency, EMMA_health_consequence, EMMA_verification, EMMA_begin_date, EMMA_end_date)
	call navigate_to_screen("STAT", "EMMA")
	call ERRR_screen_check
	call create_panel_if_nonexistent
	EMWriteScreen EMMA_medical_emergency, 6, 46
	EMWriteScreen EMMA_health_consequence, 8, 46
	EMWriteScreen EMMA_verification, 10, 46
	call create_MAXIS_friendly_date(EMMA_begin_date, 0, 12, 46)
	call create_MAXIS_friendly_date(EMMA_end_date, 0, 14, 46)
End function

'call write_panel_to_MAXIS_EMMA(EMMA_medical_emergency, EMMA_health_consequence, EMMA_verification, EMMA_begin_date, EMMA_end_date)

'STIN----------------------------------------------------------------------------------------------------------------------------------

STIN_type_1 = "01"
STIN_amt_1 = "1000"
STIN_avail_date_1 = "01/01/2015"
STIN_months_covered_1 = "01/15-01/16"
STIN_ver_1 = "1"
STIN_type_2 = "02"
STIN_amt_2 = "2000"
STIN_avail_date_2 = "01/02/2015"
STIN_months_covered_2 = "02/15-02/16"
STIN_ver_2 = "2"

Function write_panel_to_MAXIS_STIN(STIN_type_1, STIN_amt_1, STIN_avail_date_1, STIN_months_covered_1, STIN_ver_1, STIN_type_2, STIN_amt_2, STIN_avail_date_2, STIN_months_covered_2, STIN_ver_2)
	call navigate_to_screen("STAT", "STIN")
	call ERRR_screen_check
	call create_panel_if_nonexistent
	EMWriteScreen STIN_type_1, 8, 27				'STIN 1
	EMWriteScreen STIN_amt_1, 8, 34
	call create_MAXIS_friendly_date(STIN_avail_date_1, 0, 8, 46)
	EMWriteScreen left(STIN_months_covered_1, 2), 8, 58
	EMWriteScreen mid(STIN_months_covered_1, 4, 2), 8, 61
	EMWriteScreen mid(STIN_months_covered_1, 7, 2), 8, 67
	EMWriteScreen right(STIN_months_covered_1, 2), 8, 70
	EMWriteScreen STIN_ver_1, 8, 76
	EMWriteScreen STIN_type_2, 9, 27				'STIN 2
	EMWriteScreen STIN_amt_2, 9, 34
	call create_MAXIS_friendly_date(STIN_avail_date_2, 0, 9, 46)
	EMWriteScreen left(STIN_months_covered_2, 2), 9, 58
	EMWriteScreen mid(STIN_months_covered_2, 4, 2), 9, 61
	EMWriteScreen mid(STIN_months_covered_2, 7, 2), 9, 67
	EMWriteScreen right(STIN_months_covered_2, 2), 9, 70
	EMWriteScreen STIN_ver_2, 9, 76
End function

'call write_panel_to_MAXIS_STIN(STIN_type_1, STIN_amt_1, STIN_avail_date_1, STIN_months_covered_1, STIN_ver_1, STIN_type_2, STIN_amt_2, STIN_avail_date_2, STIN_months_covered_2, STIN_ver_2)

'STEC----------------------------------------------------------------------------------------------------------------------------------

STEC_type_1 = "01"
STEC_amt_1 = "1000"
STEC_actual_from_thru_months_1 = "01/15-01/16"
STEC_ver_1 = "1"
STEC_earmarked_amt_1 = "100"
STEC_earmarked_from_thru_months_1 = "01/15-01/16"
STEC_type_2 = "02"
STEC_amt_2 = "2000"
STEC_actual_from_thru_months_2 = "02/15-02/16"
STEC_ver_2 = "2"
STEC_earmarked_amt_2 = "200"
STEC_earmarked_from_thru_months_2 = "02/15-02/16"

Function write_panel_to_MAXIS_STEC(STEC_type_1, STEC_amt_1, STEC_actual_from_thru_months_1, STEC_ver_1, STEC_earmarked_amt_1, STEC_earmarked_from_thru_months_1, STEC_type_2, STEC_amt_2, STEC_actual_from_thru_months_2, STEC_ver_2, STEC_earmarked_amt_2, STEC_earmarked_from_thru_months_2)
	call navigate_to_screen("STAT", "STEC")
	call ERRR_screen_check
	call create_panel_if_nonexistent
	EMWriteScreen STEC_type_1, 8, 25				'STEC 1
	EMWriteScreen STEC_amt_1, 8, 31
	EMWriteScreen left(STEC_actual_from_thru_months_1, 2), 8, 41
	EMWriteScreen mid(STEC_actual_from_thru_months_1, 4, 2), 8, 44
	EMWriteScreen mid(STEC_actual_from_thru_months_1, 7, 2), 8, 48
	EMWriteScreen right(STEC_actual_from_thru_months_1, 2), 8, 51
	EMWriteScreen STEC_ver_1, 8, 55
	EMWriteScreen STEC_earmarked_amt_1, 8, 59
	EMWriteScreen left(STEC_earmarked_from_thru_months_1, 2), 8, 69
	EMWriteScreen mid(STEC_earmarked_from_thru_months_1, 4, 2), 8, 72
	EMWriteScreen mid(STEC_earmarked_from_thru_months_1, 7, 2), 8, 76
	EMWriteScreen right(STEC_earmarked_from_thru_months_1, 2), 8, 79
	EMWriteScreen STEC_type_2, 9, 25				'STEC 1
	EMWriteScreen STEC_amt_2, 9, 31
	EMWriteScreen left(STEC_actual_from_thru_months_2, 2), 9, 41
	EMWriteScreen mid(STEC_actual_from_thru_months_2, 4, 2), 9, 44
	EMWriteScreen mid(STEC_actual_from_thru_months_2, 7, 2), 9, 48
	EMWriteScreen right(STEC_actual_from_thru_months_2, 2), 9, 51
	EMWriteScreen STEC_ver_2, 9, 55
	EMWriteScreen STEC_earmarked_amt_2, 9, 59
	EMWriteScreen left(STEC_earmarked_from_thru_months_2, 2), 9, 69
	EMWriteScreen mid(STEC_earmarked_from_thru_months_2, 4, 2), 9, 72
	EMWriteScreen mid(STEC_earmarked_from_thru_months_2, 7, 2), 9, 76
	EMWriteScreen right(STEC_earmarked_from_thru_months_2, 2), 9, 79
End function

'call write_panel_to_MAXIS_STEC(STEC_type_1, STEC_amt_1, STEC_actual_from_thru_months_1, STEC_ver_1, STEC_earmarked_amt_1, STEC_earmarked_from_thru_months_1, STEC_type_2, STEC_amt_2, STEC_actual_from_thru_months_2, STEC_ver_2, STEC_earmarked_amt_2, STEC_earmarked_from_thru_months_2)

'SCHL----------------------------------------------------------------------------------------------------------------------------------
SCHL_status = "F"
SCHL_ver = "SC"
SCHL_type = "03"
SCHL_district_nbr = "11"
SCHL_kindergarten_start_date = ""
SCHL_grad_date = "01/15"
SCHL_grad_date_ver = "NO"
SCHL_primary_secondary_funding = "1"
SCHL_FS_eligibility_status = "03"
SCHL_higher_ed = "N"

Function write_panel_to_MAXIS_SCHL(SCHL_status, SCHL_ver, SCHL_type, SCHL_district_nbr, SCHL_kindergarten_start_date, SCHL_grad_date, SCHL_grad_date_ver, SCHL_primary_secondary_funding, SCHL_FS_eligibility_status, SCHL_higher_ed)
	call navigate_to_screen("STAT", "SCHL")
	call ERRR_screen_check
	call create_panel_if_nonexistent
	call create_MAXIS_friendly_date(date, 0, 5, 40)						'Writes actual date, needs to add 2000 as this is weirdly a 4 digit year
	EMWriteScreen datepart("yyyy", date), 5, 46
	EMWriteScreen SCHL_status, 6, 40
	EMWriteScreen SCHL_ver, 6, 63
	EMWriteScreen SCHL_type, 7, 40
	EMWriteScreen SCHL_district_nbr, 8, 40
	If SCHL_kindergarten_start_date <> "" then call create_MAXIS_friendly_date(SCHL_kindergarten_start_date, 0, 10, 63)
	EMWriteScreen left(SCHL_grad_date, 2), 11, 63
	EMWriteScreen right(SCHL_grad_date, 2), 11, 66
	EMWriteScreen SCHL_grad_date_ver, 12, 63
	EMWriteScreen SCHL_primary_secondary_funding, 14, 63
	EMWriteScreen SCHL_FS_eligibility_status, 16, 63
	EMWriteScreen SCHL_higher_ed, 18, 63
	transmit
	transmit
End function

'call write_panel_to_MAXIS_SCHL(SCHL_status, SCHL_ver, SCHL_type, SCHL_district_nbr, SCHL_kindergarten_start_date, SCHL_grad_date, SCHL_grad_date_ver, SCHL_primary_secondary_funding, SCHL_FS_eligibility_status, SCHL_higher_ed)

'MEDI----------------------------------------------------------------------------------------------------------------------------------
MEDI_claim_number_suffix = "A"
MEDI_part_A_premium = "0"
MEDI_part_B_premium = "104.90"
MEDI_part_A_begin_date = "01/01/2014"
MEDI_part_B_begin_date = "01/01/2014"

Function write_panel_to_MAXIS_MEDI(SSN_first, SSN_mid, SSN_last, MEDI_claim_number_suffix, MEDI_part_A_premium, MEDI_part_B_premium, MEDI_part_A_begin_date, MEDI_part_B_begin_date)
	call navigate_to_screen("STAT", "MEDI")
	call ERRR_screen_check
	call create_panel_if_nonexistent
	EMWriteScreen SSN_first, 6, 44				'Next three lines pulled
	EMWriteScreen SSN_mid, 6, 48
	EMWriteScreen SSN_last, 6, 51
	EMWriteScreen MEDI_claim_number_suffix, 6, 56
	EMWriteScreen MEDI_part_A_premium, 7, 46
	EMWriteScreen MEDI_part_B_premium, 7, 73
	If MEDI_part_A_begin_date <> "" then call create_MAXIS_friendly_date(MEDI_part_A_begin_date, 0, 15, 24)
	If MEDI_part_B_begin_date <> "" then call create_MAXIS_friendly_date(MEDI_part_B_begin_date, 0, 15, 54)
	transmit
	transmit
End function

'call write_panel_to_MAXIS_MEDI(MEDI_claim_number_suffix, MEDI_part_A_premium, MEDI_part_B_premium, MEDI_part_A_begin_date, MEDI_part_B_begin_date)

'FACI----------------------------------------------------------------------------------------------------------------------------------

FACI_vendor_number = ""
FACI_name = "The dumb faci"
FACI_type = "41"
FACI_FS_eligible = "N"
FACI_FS_facility_type = "6"
FACI_date_in = "11/14/2013"
FACI_date_out = "12/14/2013"

Function write_panel_to_MAXIS_FACI(FACI_vendor_number, FACI_name, FACI_type, FACI_FS_eligible, FACI_FS_facility_type, FACI_date_in, FACI_date_out)
	call navigate_to_screen("STAT", "FACI")
	call ERRR_screen_check
	call create_panel_if_nonexistent
	EMWriteScreen FACI_vendor_number, 5, 43
	EMWriteScreen FACI_name, 6, 43
	EMWriteScreen FACI_type, 7, 43
	EMWriteScreen FACI_FS_eligible, 8, 43
	If FACI_date_in <> "" then 
		call create_MAXIS_friendly_date(FACI_date_in, 0, 14, 47)
		EMWriteScreen datepart("YYYY", FACI_date_in), 14, 53
	End if
	If FACI_date_out <> "" then 
		call create_MAXIS_friendly_date(FACI_date_out, 0, 14, 71)
		EMWriteScreen datepart("YYYY", FACI_date_out), 14, 77
	End if
	transmit
	transmit
End function

'call write_panel_to_MAXIS_FACI(FACI_vendor_number, FACI_name, FACI_type, FACI_FS_eligible, FACI_FS_facility_type, FACI_date_in, FACI_date_out)