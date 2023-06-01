'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - EX PARTE REPORT.vbs"
start_time = timer
STATS_counter = 1			     'sets the stats counter at one
STATS_manualtime = 	100			 'manual run time in seconds
STATS_denomination = "C"		 'C is for each case
'END OF stats block==============================================================================================

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
call changelog_update("05/10/2023", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'FUNCTIONS =================================================================================================================

function find_unea_information()
	Call navigate_to_MAXIS_screen("STAT", "UNEA")
	For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
		EMWriteScreen MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), 20, 76
		EMWriteScreen "01", 20, 79
		transmit
		MEMBER_INFO_ARRAY(unea_VA_exists, each_memb) = False
		MEMBER_INFO_ARRAY(unea_UC_exists, each_memb) = False

		EMReadScreen unea_vers, 1, 2, 78
		If unea_vers <> "0" Then
			Do
				EMReadScreen claim_num, 15, 6, 37
				EMReadScreen income_type_code, 2, 5, 37
				If income_type_code = "01" or income_type_code = "02" Then
					If left(start_of_claim, 9) <> MEMBER_INFO_ARRAY(memb_ssn_const, each_memb) Then
						MEMBER_INFO_ARRAY(unmatched_claim_numb, each_memb) = claim_num
					End If
				End if
				claim_num = replace(claim_num, "_", "")

				If income_type_code = "11" or income_type_code = "12" or income_type_code = "13" or income_type_code = "38" Then
					MEMBER_INFO_ARRAY(unea_VA_exists, each_memb) = True
					ReDim Preserve VA_INCOME_ARRAY(va_last_const, va_count)

					VA_INCOME_ARRAY(va_case_numb_const, va_count) = MAXIS_case_number
					VA_INCOME_ARRAY(va_ref_numb_const, va_count) = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
					VA_INCOME_ARRAY(va_pers_name_const, va_count) = MEMBER_INFO_ARRAY(memb_name_const, each_memb)
					VA_INCOME_ARRAY(va_pers_ssn_const, va_count) = left(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 3) & "-" & mid(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4, 2) & "-" & right(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4)
					VA_INCOME_ARRAY(va_pers_pmi_const, va_count) = MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb)
					VA_INCOME_ARRAY(va_inc_type_code_const, va_count) = income_type_code
					If income_type_code = "11" Then VA_INCOME_ARRAY(va_inc_type_info_const, va_count) = "VA Disability"
					If income_type_code = "12" Then VA_INCOME_ARRAY(va_inc_type_info_const, va_count) = "VA Pension"
					If income_type_code = "13" Then VA_INCOME_ARRAY(va_inc_type_info_const, va_count) = "VA Other"
					If income_type_code = "38" Then VA_INCOME_ARRAY(va_inc_type_info_const, va_count) = "VA Aid & Attendance"
					VA_INCOME_ARRAY(va_claim_numb_const, va_count) = claim_num
					EMReadScreen VA_INCOME_ARRAY(va_prosp_inc_const, va_count), 8, 18, 68
					VA_INCOME_ARRAY(va_prosp_inc_const, va_count) = trim(VA_INCOME_ARRAY(va_prosp_inc_const, va_count))
					If VA_INCOME_ARRAY(va_prosp_inc_const, va_count) = "" Then VA_INCOME_ARRAY(va_prosp_inc_const, va_count) = "0.00"

					va_count = va_count + 1
				End If

				If income_type_code = "14" Then
					MEMBER_INFO_ARRAY(unea_UC_exists, each_memb) = True
					ReDim Preserve UC_INCOME_ARRAY(uc_last_const, uc_count)

					UC_INCOME_ARRAY(uc_case_numb_const, uc_count) = MAXIS_case_number
					UC_INCOME_ARRAY(uc_ref_numb_const, uc_count) = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
					UC_INCOME_ARRAY(uc_pers_name_const, uc_count) = MEMBER_INFO_ARRAY(memb_name_const, each_memb)
					UC_INCOME_ARRAY(uc_pers_ssn_const, uc_count) = left(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 3) & "-" & mid(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4, 2) & "-" & right(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4)
					UC_INCOME_ARRAY(uc_pers_pmi_const, uc_count) = MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb)
					UC_INCOME_ARRAY(uc_inc_type_code_const, uc_count) = income_type_code
					UC_INCOME_ARRAY(uc_inc_type_info_const, uc_count) = "Unemployment"
					UC_INCOME_ARRAY(uc_claim_numb_const, uc_count) = claim_num
					EMReadScreen UC_INCOME_ARRAY(uc_prosp_inc_const, uc_count), 8, 13, 68
					UC_INCOME_ARRAY(uc_prosp_inc_const, uc_count) = trim(UC_INCOME_ARRAY(uc_prosp_inc_const, uc_count))
					If UC_INCOME_ARRAY(uc_prosp_inc_const, uc_count) = "________" Then UC_INCOME_ARRAY(uc_prosp_inc_const, uc_count) = "0.00"

					uc_count = uc_count + 1
				End If

				transmit
				EMReadScreen next_unea_nav, 7, 24, 2
			Loop until next_unea_nav = "ENTER A"
		End If
	Next
end function

function find_UNEA_panel(MEMB_reference_number, UNEA_type_code, UNEA_instance, UNEA_claim_number, panel_found)
	panel_found = False
	UNEA_claim_number = replace(UNEA_claim_number, " ", "")
	EMWriteScreen "UNEA", 20, 71
	transmit
	EMReadScreen unea_check, 4, 2, 48
	Do While unea_check <> "UNEA"
		Call navigate_to_MAXIS_screen("STAT", "UNEA")
		EMReadScreen unea_check, 4, 2, 48
	Loop
	EMWriteScreen MEMB_reference_number, 20, 76 		'Navigating to STAT/UNEA
	EMWriteScreen "01", 20, 79 		'to ensure we're on the 1st instance of UNEA panels for the appropriate member
	transmit

	EMReadScreen vers_count, 1, 2, 78
	If vers_count <> "0" Then
		Do
			EMReadScreen panel_type_code, 2, 5, 37
			EMReadScreen panel_claim_number, 15, 6, 37

			panel_claim_number = replace(panel_claim_number, "_", "")
			panel_claim_number = replace(panel_claim_number, " ", "")

			If panel_type_code = UNEA_type_code Then panel_found = True

			If UNEA_type_code = "01" or UNEA_type_code = "02" Then
				If UNEA_claim_number = panel_claim_number Then panel_found = True
			End If
			' MsgBox "panel_type_code - " & panel_type_code & vbcr & "UNEA_type_code - " & UNEA_type_code & vbCr & vbCr &_
			' 		"panel_claim_number - " & panel_claim_number & vbCr & "UNEA_claim_number - " & UNEA_claim_number & vbCr & vbCr &_
			' 		"panel_found - " & panel_found
			If panel_found = True Then
				EMReadScreen UNEA_instance, 1, 2, 73
				UNEA_instance = "0" & UNEA_instance
				Exit Do
			End If

			transmit
			EMReadScreen end_of_UNEA_panels, 7, 24, 2
		Loop Until end_of_UNEA_panels = "ENTER A"
	End If
end function

function get_list_of_members()
	client_count = UBound(MEMBER_INFO_ARRAY, 2) + 1
	If MEMBER_INFO_ARRAY(memb_pmi_numb_const, 0) = "" Then client_count = 0
	EMWriteScreen "01", 20, 76						''make sure to start at Memb 01
	transmit
	DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
		EMReadScreen client_PMI, 8, 4, 46
		client_PMI = trim(client_PMI)
		client_PMI = RIGHT("00000000" & client_PMI, 8)

		client_found = False
		For known_membs = 0 to UBound(MEMBER_INFO_ARRAY, 2)
			If client_PMI = MEMBER_INFO_ARRAY(memb_pmi_numb_const, known_membs) Then
				client_found = True
				EMReadScreen MEMBER_INFO_ARRAY(memb_ref_numb_const, known_membs), 2, 4, 33
				EMReadScreen clt_age, 3, 8, 76
				MEMBER_INFO_ARRAY(memb_age_const, known_membs) = trim(clt_age)
				Exit For
			End If
		Next

		If client_found = False Then
			ReDim Preserve MEMBER_INFO_ARRAY(memb_last_const, client_count)

			EMReadScreen MEMBER_INFO_ARRAY(memb_ref_numb_const, client_count), 2, 4, 33
			MEMBER_INFO_ARRAY(memb_pmi_numb_const, client_count) = client_PMI
			EMReadScreen SSN1, 3, 7, 42
			EMReadScreen SSN2, 2, 7, 46
			EMReadScreen SSN3, 4, 7, 49
			MEMBER_INFO_ARRAY(memb_ssn_const, client_count) = SSN1 & SSN2 & SSN3
			EMReadScreen clt_age, 3, 8, 76
			MEMBER_INFO_ARRAY(memb_age_const, client_count) = trim(clt_age)
			EMReadScreen last_name, 25, 6, 30
			EMReadScreen first_name, 12, 6, 63
			last_name = trim(replace(last_name, "_", ""))
			first_name = trim(replace(first_name, "_", ""))
			MEMBER_INFO_ARRAY(memb_name_const, client_count) = last_name & ", " & first_name
			MEMBER_INFO_ARRAY(memb_active_hc_const, client_count)	= False

			client_count = client_count + 1
		End If
		transmit
		EMReadScreen edit_check, 7, 24, 2
	LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.
end function

function send_sves_qury(ssn_or_claim, qury_finish)
	qury_finish = ""
	Call navigate_to_MAXIS_screen("INFC", "SVES")
	EMWriteScreen MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4, 68
	EMWriteScreen MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb), 5, 68
	Call write_value_and_transmit("QURY", 20, 70)										'Now we will enter the QURY screen to type the case number.

	If ssn_or_claim = "CLAIM" Then
		Call clear_line_of_text(5, 38)
		EMWriteScreen MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 7, 38
	End If
	EMWriteScreen MAXIS_case_number, 	11, 38
	EMWriteScreen "Y", 					14, 38
	transmit  'Now it sends the SVES.

	EMReadScreen duplicate_SVES, 	    7, 24, 2
	If duplicate_SVES = "WARNING" then transmit
	EMReadScreen confirm_SVES, 			6, 24, 2
	if confirm_SVES = "RECORD" then
		' PMI_array(SVES_status, item) = True
		qury_finish = date
	Else
		' PMI_array(SVES_status, item) = False
		qury_finish = "FAILED"
	END IF
end function

function update_unea_pane(panel_found, unea_type, income_amount, claim_number, start_date, end_date)
	panel_in_edit_mode = False
	If panel_found = False and ssi_end_date = "" Then
		Call write_value_and_transmit("NN", 20, 79)
		panel_in_edit_mode = True
	ElseIf panel_found = True Then
		PF9
		panel_in_edit_mode = True
	End If
	If panel_in_edit_mode = True Then
		Call clear_line_of_text(6, 37)
		EMWriteScreen claim_number, 6, 37
		EMWriteScreen unea_type, 5, 37
		EMWriteScreen "7", 5, 65		'Write Verification Worker Initiated Verfication "7"
		If panel_found = False Then Call create_mainframe_friendly_date(start_date, 7, 37, "YY") 	'income start date (SSI: ssi_SSP_elig_date, RSDI: intl_entl_date)

		Call clear_line_of_text(10, 67)		'clear the COLA disregard - TODO - update this for Jan - June to not remove this
		Call clear_line_of_text(7, 68)
		Call clear_line_of_text(7, 71)
		Call clear_line_of_text(7, 74)
		'Clear amounts
		row = 13
		DO
			EMWriteScreen "__", row, 25
			EMWriteScreen "__", row, 28
			EMWriteScreen "__", row, 31
			EMWriteScreen "________", row, 39

			EMWriteScreen "__", row, 54
			EMWriteScreen "__", row, 57
			EMWriteScreen "__", row, 60
			EMWriteScreen "________", row, 68
			row = row + 1
		Loop until row = 18

		EMWriteScreen CM_minus_1_mo, 13, 25 'hardcoded dates
		EMWriteScreen "01", 13, 28
		EMWriteScreen CM_minus_1_yr, 13, 31 'hardcoded dates
		EMWriteScreen income_amount, 13, 39		'TODO: Testing values

		EMWriteScreen CM_plus_1_mo, 13, 54 'hardcoded dates
		EMWriteScreen "01", 13, 57
		EMWriteScreen CM_plus_1_yr, 13, 60 'hardcoded dates
		EMWriteScreen income_amount, 13, 68		'TODO: Testing values (income_amt which = rsdi_gross_amt or ssi_gross_amt )
		If end_date <> "" Then Call create_mainframe_friendly_date(end_date, 7, 68, "YY")	'income end date (SSI: ssi_denial_date, RSDI: susp_term_date)

		Call write_value_and_transmit("X", 6, 56)
		Call clear_line_of_text(9, 65)
		EMWriteScreen income_amount, 9, 65		'TODO: Testing values (rsdi_gross_amt or ssi_gross_amt )
		EMWriteScreen "1", 10, 63		'code for pay frequency
		Do
			transmit
			EMReadScreen HC_popup, 9, 7, 41
			' If HC_popup = "HC Income" then transmit
		Loop until HC_popup <> "HC Income"

		' MsgBox "STOP AND LOOK AT THE PANEL"
		' PF10

		transmit
		EMReadScreen cola_warning, 29, 24, 2
		If cola_warning = "WARNING: ENTER COLA DISREGARD" then transmit
		EMReadScreen HC_income_warning, 25, 24, 2
		If HC_income_warning = "WARNING: UPDATE HC INCOME" then transmit
		' MsgBox "Wait"
	End If
end function


'END FUNCTIONS BLOCK =======================================================================================================


'DECLARATIONS ==============================================================================================================

Const memb_ref_numb_const 	= 0
Const memb_pmi_numb_const 	= 1
Const memb_ssn_const 		= 2
Const memb_age_const 		= 3
Const memb_name_const 		= 4
Const memb_active_hc_const	= 5
Const table_prog_1			= 6
Const table_type_1			= 7
Const table_prog_2			= 8
Const table_type_2			= 9
Const table_prog_3			= 10
Const table_type_3			= 11

Const unea_type_01_esists	= 20
Const unea_type_02_esists	= 21
Const unea_type_03_esists	= 22
Const unea_type_16_esists	= 23
Const unmatched_claim_numb	= 24
Const unea_VA_exists		= 25
Const unea_UC_exists		= 26

Const sves_qury_sent		= 40
Const second_qury_sent		= 41
Const sves_tpqy_response	= 42

Const tpqy_rsdi_record 				= 45
Const tpqy_ssi_record 				= 46
Const tpqy_rsdi_claim_numb 			= 47
Const tpqy_dual_entl_nbr 			= 48
Const tpqy_rsdi_status_code 		= 49
Const tpqy_rsdi_gross_amt 			= 50
Const tpqy_rsdi_net_amt 			= 51
Const tpqy_railroad_ind 			= 52
Const tpqy_intl_entl_date 			= 53
Const tpqy_susp_term_date 			= 54
Const tpqy_rsdi_disa_date 			= 55
Const tpqy_medi_claim_num 			= 56
Const tpqy_part_a_premium 			= 57
Const tpqy_part_a_start 			= 58
Const tpqy_part_a_stop 				= 59
Const tpqy_part_a_buyin_ind 		= 60
Const tpqy_part_a_buyin_code 		= 61
Const tpqy_part_a_buyin_start_date 	= 62
Const tpqy_part_a_buyin_stop_date 	= 63
Const tpqy_part_b_premium 			= 64
Const tpqy_part_b_start 			= 65
Const tpqy_part_b_stop 				= 66
Const tpqy_part_b_buyin_ind 		= 67
Const tpqy_Part_b_buyin_code 		= 68
Const tpqy_part_b_buyin_start_date 	= 69
Const tpqy_part_b_buyin_stop_date 	= 70
Const tpqy_ssi_claim_numb 			= 71
Const tpqy_ssi_recip_code 			= 72
Const tpqy_ssi_recip_desc 			= 73
Const tpqy_fed_living 				= 74
Const tpqy_ssi_pay_code 			= 75
Const tpqy_ssi_pay_desc 			= 76
Const tpqy_cit_ind_code 			= 77
Const tpqy_ssi_denial_code 			= 78
Const tpqy_ssi_denial_desc 			= 79
Const tpqy_ssi_denial_date 			= 80
Const tpqy_ssi_disa_date 			= 81
Const tpqy_ssi_SSP_elig_date 		= 82
Const tpqy_ssi_appeals_code 		= 83
Const tpqy_ssi_appeals_date 		= 84
Const tpqy_ssi_appeals_dec_code 	= 85
Const tpqy_ssi_appeals_dec_date 	= 86
Const tpqy_ssi_disa_pay_code 		= 87
Const tpqy_ssi_pay_date 			= 88
Const tpqy_ssi_gross_amt 			= 89
Const tpqy_ssi_over_under_code 		= 90
Const tpqy_ssi_pay_hist_1_date 		= 91
Const tpqy_ssi_pay_hist_1_amt 		= 92
Const tpqy_ssi_pay_hist_1_type 		= 93
Const tpqy_ssi_pay_hist_2_date 		= 94
Const tpqy_ssi_pay_hist_2_amt 		= 95
Const tpqy_ssi_pay_hist_2_type 		= 96
Const tpqy_ssi_pay_hist_3_date 		= 97
Const tpqy_ssi_pay_hist_3_amt 		= 98
Const tpqy_ssi_pay_hist_3_type 		= 99
Const tpqy_gross_EI 				= 100
Const tpqy_net_EI 					= 101
Const tpqy_rsdi_income_amt 			= 102
Const tpqy_pass_exclusion 			= 103
Const tpqy_inc_inkind_start 		= 104
Const tpqy_inc_inkind_stop 			= 105
Const tpqy_rep_payee 				= 106

Const tpqy_memb_has_ssi				= 110
Const tpqy_memb_has_rsdi			= 111
Const tpqy_rsdi_has_disa			= 112
Const created_medi					= 113
Const updated_medi_a				= 114
Const updated_medi_b				= 115


Const memb_last_const 		= 120

Dim MEMBER_INFO_ARRAY()


Const va_case_numb_const 	= 0
Const va_ref_numb_const 	= 1
Const va_pers_name_const	= 2
Const va_pers_pmi_const		= 3
Const va_pers_ssn_const		= 4
Const va_inc_type_code_const 	= 5
Const va_inc_type_info_const	= 6
Const va_claim_numb_const 	= 7
Const va_prosp_inc_const 	= 8
Const va_last_const 		= 9

Dim VA_INCOME_ARRAY()
ReDim VA_INCOME_ARRAY(va_last_const, 0)

Const uc_case_numb_const 	= 0
Const uc_ref_numb_const 	= 1
Const uc_pers_name_const	= 2
Const uc_pers_pmi_const		= 3
Const uc_pers_ssn_const		= 4
Const uc_inc_type_code_const 	= 5
Const uc_inc_type_info_const	= 6
Const uc_claim_numb_const 	= 7
Const uc_prosp_inc_const 	= 8
Const uc_last_const 		= 9

Dim UC_INCOME_ARRAY()
ReDim UC_INCOME_ARRAY(va_last_const, 0)


'END DECLARATIONS BLOCK ====================================================================================================






'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone

Confirm_Process_to_Run_btn	= 200
incorrect_process_btn		= 100
end_msg = "DONE"

If Day(date) < 1 Then ex_parte_function = "Prep"

'DISPLAYS DIALOG

DO
	DO
		DO
			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 401, 255, "Ex Parte Report"
				DropListBox 300, 25, 90, 15, "Select one..."+chr(9)+"Prep"+chr(9)+"Phase 1"+chr(9)+"Phase 2"+chr(9)+"FIX LIST", ex_parte_function
				ButtonGroup ButtonPressed
					OkButton 290, 235, 50, 15
					CancelButton 345, 235, 50, 15
				Text 5, 10, 400, 10, "This script will connect to the SQL Table to pull a list of cases to operate on based on the Ex Parte functionality selected."
				Text 200, 30, 95, 10, "Selection Ex Parte Function:"
				Text 10, 45, 35, 10, "Prep"
				Text 50, 45, 150, 10, "Timing - 4 Days before the BUDGET Month"
				Text 50, 55, 190, 10, "Collect any Case Criteria not available in Info Store."
				Text 50, 65, 175, 10, "Send SVES/QURY for all members on all cases."
				Text 50, 75, 200, 10, "Generate a UC and VA Verif Report for OS Staff completion."
				Text 10, 90, 35, 10, "Phase 1"
				Text 50, 90, 135, 10, "Timing - 1st Day of the BUDGET Month"
				Text 50, 100, 245, 10, "Read SVES/TPQY Response, Update STAT with detail, enter CASE/NOTE."
				Text 50, 110, 270, 10, "Udate STAT with UC or VA Verifications provided from OS Report and CASE/NOTE."
				Text 50, 120, 125, 10, "Run each case through Background."
				Text 50, 130, 200, 10, "Read and Record in the SQL Table the ELIG information."
				Text 50, 140, 225, 10, "Read and Record in the SQL Table the detail of MMIS Open Spans."
				Text 10, 155, 35, 10, "Phase 2"
				Text 50, 155, 160, 10, "Timing - 1st Day of the PROCESSING Month"
				Text 50, 165, 285, 10, "Check DAIL, CASE/NOTE, STAT for any updates since Phase 1 Ex Parte Determination."
				Text 50, 175, 145, 10, "Record in SQL Table any Updates found."
				Text 50, 185, 125, 10, "Run each case through Background."
				Text 50, 195, 200, 10, "Read and Record in the SQL Table the ELIG information."
				Text 10, 215, 205, 10, "* * * * * THIS SCRIPT MUST BE RUN IN PRODUCTION * * * * *"
				Text 10, 235, 190, 10, "There is no CASE/NOTE entry by this script at this time."
			EndDialog

			err_msg = ""
			Dialog Dialog1
			cancel_without_confirmation
			If ex_parte_function = "Select one..." then err_msg = err_msg & vbNewLine & "* Select an Ex Parte Function."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP until err_msg = ""

		If ex_parte_function = "Prep" or ex_parte_function = "FIX LIST" Then
			ep_revw_mo = right("00" & DatePart("m",	DateAdd("m", 3, date)), 2)
			ep_revw_yr = right(DatePart("yyyy",	DateAdd("m", 3, date)), 2)
			' ep_revw_mo = "07"
			' ep_revw_yr = "23"

		End If
		If ex_parte_function = "Phase 1" Then
			ep_revw_mo = right("00" & DatePart("m",	DateAdd("m", 2, date)), 2)
			ep_revw_yr = right(DatePart("yyyy",	DateAdd("m", 2, date)), 2)
			ep_revw_mo = "08"
			ep_revw_yr = "23"
		End If
		If ex_parte_function = "Phase 2" Then
			ep_revw_mo = right("00" & DatePart("m",	DateAdd("m", 1, date)), 2)
			ep_revw_yr = right(DatePart("yyyy",	DateAdd("m", 1, date)), 2)

		End If

		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 341, 165, "Confirm Ex Parte process"
			EditBox 600, 700, 10, 10, fake_edit_box
			ButtonGroup ButtonPressed
				PushButton 10, 145, 210, 15, "CONFIRMED! This is the correct Process and Review Month", Confirm_Process_to_Run_btn
				PushButton 230, 145, 100, 15, "Incorrect Process/Month", incorrect_process_btn
			Text 10, 10, 225, 10, "You are running the Ex Parte Function " & ex_parte_function
			Text 10, 25, 190, 10, "This will run for the Ex Parte Review month of " & ep_revw_mo & "/" & ep_revw_yr
			If ex_parte_function = "Prep" or ex_parte_function = "FIX LIST" Then
				GroupBox 5, 40, 240, 50, "Tasks to be Completed:"
				Text 20, 55, 190, 10, "Collect any Case Criteria not available in Info Store."
				Text 20, 65, 175, 10, "Send SVES/QURY for all members on all cases."
				Text 20, 75, 200, 10, "Generate a UC and VA Verif Report for OS Staff completion."
			End If
			If ex_parte_function = "Phase 1" Then
				GroupBox 5, 40, 295, 70, "Tasks to be Completed:"
				Text 20, 55, 245, 10, "Read SVES/TPQY Response, Update STAT with detail, enter CASE/NOTE."
				Text 20, 65, 270, 10, "Udate STAT with UC or VA Verifications provided from OS Report and CASE/NOTE."
				Text 20, 75, 125, 10, "Run each case through Background."
				Text 20, 85, 200, 10, "Read and Record in the SQL Table the ELIG information."
				Text 20, 95, 225, 10, "Read and Record in the SQL Table the detail of MMIS Open Spans."
			End If
			If ex_parte_function = "Phase 2" Then
				GroupBox 5, 40, 305, 60, "Tasks to be Completed:"
				Text 20, 55, 285, 10, "Check DAIL, CASE/NOTE, STAT for any updates since Phase 1 Ex Parte Determination."
				Text 20, 65, 145, 10, "Record in SQL Table any Updates found."
				Text 20, 75, 125, 10, "Run each case through Background."
				Text 20, 85, 200, 10, "Read and Record in the SQL Table the ELIG information."
			End If
			Text 10, 115, 190, 10, "There is no CASE/NOTE entry by this script at this time."
			Text 10, 130, 330, 10, "Review the process datails and ex parte review month to confirm this is the correct run to complete."
		EndDialog

		Dialog Dialog1
		cancel_without_confirmation

		If ButtonPressed = OK Then ButtonPressed = Confirm_Process_to_Run_btn

	Loop until ButtonPressed = Confirm_Process_to_Run_btn
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

If ex_parte_function = "FIX LIST" Then
	Call script_end_procedure("There is no fix currently established.")
	prep_status = "Not Ex Parte"
	appears_ex_parte = False
	MAXIS_footer_month = CM_plus_1_mo
	MAXIS_footer_year = CM_plus_1_yr

	review_date = ep_revw_mo & "/1/" & ep_revw_yr
	review_date = DateAdd("d", 0, review_date)

	'declare the SQL statement that will query the database
	objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "' and [SelectExParte] = '1' and [NoIncome] = 'False'"
	' objSQL = "UPDATE ES.ES_ExParte_CaseList SET SelectExParte = '" & appears_ex_parte & "', PREP_Complete = '" & prep_status & "' WHERE [HCEligReviewDate] = '" & review_date & "' and [SelectExParte] = '1' and [VAIncomeExist] = 'True'"

	'Creating objects for Access
	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'This is the file path for the statistics Access database.
	' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	'Setting a starting value for a list of cases so that every case is bracketed by * on both sides.
	todays_cases_list = "*"
	case_entry = 0      'Setting an incrementor for the array to be filled

	'Open The CASE LIST Table
	'Loop through each item on the CASE LIST Table

	Do While NOT objRecordSet.Eof
		MAXIS_case_number = objRecordSet("CaseNumber") 		'SET THE MAXIS CASE NUMBER
		UC_income_exists = False

		objIncomeSQL = "SELECT * FROM ES.ES_ExParte_IncomeList WHERE [CaseNumber] = '" & MAXIS_case_number & "'"

		'Creating objects for Access
		Set objIncomeConnection = CreateObject("ADODB.Connection")
		Set objIncomeRecordSet = CreateObject("ADODB.Recordset")

		'This is the file path for the statistics Access database.
		' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
		objIncomeConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
		objIncomeRecordSet.Open objIncomeSQL, objIncomeConnection

		Do While NOT objIncomeRecordSet.Eof
			If objIncomeRecordSet("IncExpTypeCode") = "UNEA" Then
				If objIncomeRecordSet("IncomeTypeCode") = "14" Then UC_income_exists = True
			End If
			objIncomeRecordSet.MoveNext
		Loop
		objIncomeRecordSet.Close
		objIncomeConnection.Close
		Set objIncomeRecordSet=nothing
		Set objIncomeConnection=nothing

		If UC_income_exists = True Then

			objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET SelectExParte = '" & appears_ex_parte & "', PREP_Complete = '" & prep_status & "' WHERE CaseNumber = '" & MAXIS_case_number & "'"

			'Creating objects for Access
			Set objUpdateConnection = CreateObject("ADODB.Connection")
			Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

			'This is the file path for the statistics Access database.
			objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
			objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection
		End IF
		objRecordSet.MoveNext
	Loop


End If

If ex_parte_function = "Prep" Then
	' MsgBox  "At this point the script will pull the cases from a SQL Table that has identified cases due for a HC ER and evaluates them as potentially Ex Parte." & vbCr & vbCr &_
	' 		"If the case is potentially Ex Parte, the script will:" & vbCr &_
	' 		" - Send a SVES/QURY." & vbCr &_
	' 		" - Add the case to a report if VA Income is listed on the case to gather verification." & vbCr & vbCR &_
	' 		"This script will look at each case for the specified review month, preparing the case for review." & vbCr  & vbCr &_
	' 		"This script is run 4 business days before the Budget Month, or the end of the 3rd month BEFORE the ER month."

	MAXIS_footer_month = CM_plus_1_mo
	MAXIS_footer_year = CM_plus_1_yr

	review_date = ep_revw_mo & "/1/" & ep_revw_yr
	review_date = DateAdd("d", 0, review_date)

	va_count = 0
	uc_count = 0
	ex_parte_cases_count = 0

	'Open The CASE LIST Table
	'declare the SQL statement that will query the database
	objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"

	'Creating objects for Access
	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'This is the file path for the statistics Access database.
	' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	'Setting a starting value for a list of cases so that every case is bracketed by * on both sides.
	todays_cases_list = "*"
	case_entry = 0      'Setting an incrementor for the array to be filled

	'Loop through each item on the CASE LIST Table
	Do While NOT objRecordSet.Eof
		If IsNull(objRecordSet("PREP_Complete")) = True Then
			all_hc_is_ABD = ""
			SSA_income_exists = ""
			JOBS_income_exists = ""
			VA_income_exists = ""
			BUSI_income_exists = ""
			case_has_no_income = ""
			case_has_EPD = ""

			appears_ex_parte = True
			all_hc_is_ABD = True
			case_has_EPD = False
			case_is_in_henn = False
			MAXIS_case_number = objRecordSet("CaseNumber") 		'SET THE MAXIS CASE NUMBER
			' MsgBox MAXIS_case_number
			ReDim MEMBER_INFO_ARRAY(memb_last_const, 0)
			memb_count = 0

			objELIGSQL = "SELECT * FROM ES.ES_ExParte_EligList WHERE [CaseNumb] = '" & MAXIS_case_number & "'"

			'Creating objects for Access
			Set objELIGConnection = CreateObject("ADODB.Connection")
			Set objELIGRecordSet = CreateObject("ADODB.Recordset")

			'This is the file path for the statistics Access database.
			' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
			objELIGConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
			objELIGRecordSet.Open objELIGSQL, objELIGConnection

			Do While NOT objELIGRecordSet.Eof

			' If objELIGRecordSet("MajorProgram") = NULL
				memb_known = False
				For known_membs = 0 to UBound(MEMBER_INFO_ARRAY, 2)
					If trim(objELIGRecordSet("PMINumber")) = MEMBER_INFO_ARRAY(memb_pmi_numb_const, known_membs) Then
						memb_known = True
						If MEMBER_INFO_ARRAY(table_prog_1, known_membs) = "" Then
							MEMBER_INFO_ARRAY(table_prog_1, known_membs) 		= objELIGRecordSet("MajorProgram")
							MEMBER_INFO_ARRAY(table_type_1, known_membs) 		= objELIGRecordSet("EligType")
						ElseIf MEMBER_INFO_ARRAY(table_prog_2, known_membs) = "" Then
							MEMBER_INFO_ARRAY(table_prog_2, known_membs) 		= objELIGRecordSet("MajorProgram")
							MEMBER_INFO_ARRAY(table_type_2, known_membs) 		= objELIGRecordSet("EligType")
						ElseIf MEMBER_INFO_ARRAY(table_prog_3, known_membs) = "" Then
							MEMBER_INFO_ARRAY(table_prog_3, known_membs) 		= objELIGRecordSet("MajorProgram")
							MEMBER_INFO_ARRAY(table_type_3, known_membs) 		= objELIGRecordSet("EligType")
						End If
						If objELIGRecordSet("MajorProgram") = "EH" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "AX" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "AA" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "DP" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "CK" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "CX" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "CB" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "CM" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "13" Then appears_ex_parte = False 	'TYMA
						If objELIGRecordSet("EligType") = "14" Then appears_ex_parte = False 	'TYMA
						If objELIGRecordSet("EligType") = "09" Then appears_ex_parte = False 	'Adoption Assistance
						If objELIGRecordSet("EligType") = "11" Then appears_ex_parte = False 	'Auto Newborn
						If objELIGRecordSet("EligType") = "10" Then appears_ex_parte = False 	'Adoption Assistance
						If objELIGRecordSet("EligType") = "25" Then appears_ex_parte = False 	'Foster Care
						If objELIGRecordSet("EligType") = "PX" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "PC" Then appears_ex_parte = False
						If objELIGRecordSet("EligType") = "BC" Then appears_ex_parte = False

						If appears_ex_parte = False AND objELIGRecordSet("EligType") <> "DP" Then all_hc_is_ABD = False
						If objELIGRecordSet("EligType") = "DP" Then case_has_EPD = True
					End If
				Next

				If memb_known = False Then
					ReDim Preserve MEMBER_INFO_ARRAY(memb_last_const, memb_count)

					MEMBER_INFO_ARRAY(memb_pmi_numb_const, memb_count) 	= trim(objELIGRecordSet("PMINumber"))
					MEMBER_INFO_ARRAY(memb_ssn_const, memb_count) 		= trim(objELIGRecordSet("SocialSecurityNbr"))
					name_var									 		= trim(objELIGRecordSet("Name"))
					name_array = split(name_var)
					MEMBER_INFO_ARRAY(memb_name_const, memb_count) = name_array(UBound(name_array))
					For name_item = 0 to UBound(name_array)-1
						MEMBER_INFO_ARRAY(memb_name_const, memb_count) = MEMBER_INFO_ARRAY(memb_name_const, memb_count) & " " & name_array(name_item)
					Next
					MEMBER_INFO_ARRAY(memb_active_hc_const, memb_count)	= True
					MEMBER_INFO_ARRAY(table_prog_1, memb_count) 		= trim(objELIGRecordSet("MajorProgram"))
					MEMBER_INFO_ARRAY(table_type_1, memb_count) 		= trim(objELIGRecordSet("EligType"))

					If objELIGRecordSet("MajorProgram") = "EH" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "AX" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "AA" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "DP" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "CK" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "CX" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "CB" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "CM" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "13" Then appears_ex_parte = False 	'TYMA
					If objELIGRecordSet("EligType") = "14" Then appears_ex_parte = False 	'TYMA
					If objELIGRecordSet("EligType") = "09" Then appears_ex_parte = False 	'Adoption Assistance
					If objELIGRecordSet("EligType") = "11" Then appears_ex_parte = False 	'Auto Newborn
					If objELIGRecordSet("EligType") = "10" Then appears_ex_parte = False 	'Adoption Assistance
					If objELIGRecordSet("EligType") = "25" Then appears_ex_parte = False 	'Foster Care
					If objELIGRecordSet("EligType") = "PX" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "PC" Then appears_ex_parte = False
					If objELIGRecordSet("EligType") = "BC" Then appears_ex_parte = False

					If appears_ex_parte = False AND objELIGRecordSet("EligType") <> "DP" Then all_hc_is_ABD = False
					If objELIGRecordSet("EligType") = "DP" Then case_has_EPD = True

					memb_count = memb_count + 1
				End if
				objELIGRecordSet.MoveNext
			Loop
			objELIGRecordSet.Close
			objELIGConnection.Close
			Set objELIGRecordSet=nothing
			Set objELIGConnection=nothing


			SSA_income_exists = False
			RR_income_exists = False
			VA_income_exists = False
			UC_income_exists = False
			PRISM_income_exists = False
			Other_UNEA_income_exists = False
			JOBS_income_exists = False
			BUSI_income_exists = False

			objIncomeSQL = "SELECT * FROM ES.ES_ExParte_IncomeList WHERE [CaseNumber] = '" & MAXIS_case_number & "'"

			'Creating objects for Access
			Set objIncomeConnection = CreateObject("ADODB.Connection")
			Set objIncomeRecordSet = CreateObject("ADODB.Recordset")

			'This is the file path for the statistics Access database.
			' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
			objIncomeConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
			objIncomeRecordSet.Open objIncomeSQL, objIncomeConnection

			Do While NOT objIncomeRecordSet.Eof
				If objIncomeRecordSet("IncExpTypeCode") = "UNEA" Then
					If objIncomeRecordSet("IncomeTypeCode") = "01" Then SSA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "02" Then SSA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "03" Then SSA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "16" Then SSA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "11" Then VA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "12" Then VA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "13" Then VA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "38" Then VA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "14" Then UC_income_exists = True

					If objIncomeRecordSet("IncomeTypeCode") = "36" Then PRISM_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "37" Then PRISM_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "39" Then PRISM_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "40" Then PRISM_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "36" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "37" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "39" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "40" Then Other_UNEA_income_exists = True

					If objIncomeRecordSet("IncomeTypeCode") = "06" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "15" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "17" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "18" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "23" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "24" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "25" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "26" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "27" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "28" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "29" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "08" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "35" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "43" Then Other_UNEA_income_exists = True
					If objIncomeRecordSet("IncomeTypeCode") = "47" Then Other_UNEA_income_exists = True
				End If
				If objIncomeRecordSet("IncExpTypeCode") = "JOBS" Then JOBS_income_exists = True
				If objIncomeRecordSet("IncExpTypeCode") = "BUSI" Then BUSI_income_exists = True

				objIncomeRecordSet.MoveNext
			Loop
			objIncomeRecordSet.Close
			objIncomeConnection.Close
			Set objIncomeRecordSet=nothing
			Set objIncomeConnection=nothing


			If appears_ex_parte = True Then

				'check HC ER date in STAT/REVW
				Call navigate_to_MAXIS_screen_review_PRIV("STAT", "REVW", is_this_priv)
				If is_this_priv = True Then appears_ex_parte = False
				If is_this_priv = False Then
					Call write_value_and_transmit("X", 5, 71)
					EMReadScreen STAT_HC_ER_mo, 2, 8, 27
					EMReadScreen STAT_HC_ER_yr, 2, 8, 33
					If ep_revw_mo <> STAT_HC_ER_mo or ep_revw_yr <> STAT_HC_ER_yr Then  appears_ex_parte = False
				End If
			End If

			If appears_ex_parte = True Then
				Call navigate_to_MAXIS_screen("STAT", "MEMB")
				Call get_list_of_members

				Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
				EMReadScreen case_pw, 7, 21, 14
				If left(case_pw, 4) = "X127" Then case_is_in_henn = True
				If case_is_in_henn = False then  appears_ex_parte = False
				If case_active = False Then appears_ex_parte = False
				If ma_status <> "ACTIVE" and msp_status <> "ACTIVE" Then appears_ex_parte = False
				' If msp_status <> "ACTIVE" Then appears_ex_parte = False

				' If mfip_status = "ACTIVE" OR snap_status = "ACTIVE" Then
			End If

			case_has_no_income = False
			If SSA_income_exists = False and RR_income_exists = False and VA_income_exists = False and UC_income_exists = False and PRISM_income_exists = False and Other_UNEA_income_exists = False and JOBS_income_exists = False and BUSI_income_exists = False Then case_has_no_income = True

			If appears_ex_parte = True Then
				If Other_UNEA_income_exists = True OR JOBS_income_exists = True OR BUSI_income_exists = True Then
					If mfip_status = "ACTIVE" Then
						' If all_hc_is_ABD = True and case_has_EPD = False Then
						appears_ex_parte = True
					ElseIf snap_status = "ACTIVE" Then
						'find income
						appears_ex_parte = True
					Else
						' If Other_UNEA_income_exists = True Then appears_ex_parte = False
						' If JOBS_income_exists = True Then appears_ex_parte = False
						' If BUSI_income_exists = True Then appears_ex_parte = False
						appears_ex_parte = False
					End If
				End If
			End If

			If appears_ex_parte = True Then ex_parte_cases_count = ex_parte_cases_count + 1

			If appears_ex_parte = True Then
				'For each case that is indicated as potentially ExParte, we are going to take preperation actions
				last_va_count = va_count
				last_uc_count = uc_count

				Call find_unea_information

				Call back_to_SELF

				'Send a SVES/CURY for all persons on a case
				Call navigate_to_MAXIS_screen("INFC", "SVES")
				'checking for NON-DISCLOSURE AGREEMENT REQUIRED FOR ACCESS TO IEVS FUNCTIONS'
				EMReadScreen agreement_check, 9, 2, 24
				IF agreement_check = "Automated" THEN script_end_procedure("To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the screens and review the agreement.")

				'We need to loop through each HH Member on the case and send a QURY for every one.
				For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
					Call send_sves_qury("SSN", qury_finish)
					MEMBER_INFO_ARRAY(sves_qury_sent, each_memb) = qury_finish

					objIncomeSQL = "UPDATE ES.ES_ExParte_IncomeList SET QURY_Sent = '" & qury_finish & "' WHERE [CaseNumber] = '" & MAXIS_case_number & "' and [PersonID] = '" & MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb) & "' and [ClaimNbr] like '" & MEMBER_INFO_ARRAY(memb_ssn_const, each_memb) & "%'"

					'Creating objects for Access
					Set objIncomeConnection = CreateObject("ADODB.Connection")
					Set objIncomeRecordSet = CreateObject("ADODB.Recordset")

					'This is the file path for the statistics Access database.
					' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
					objIncomeConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
					objIncomeRecordSet.Open objIncomeSQL, objIncomeConnection

					MEMBER_INFO_ARRAY(second_qury_sent, each_memb) = False
					If MEMBER_INFO_ARRAY(unmatched_claim_numb, each_memb) <> "" Then
						Call send_sves_qury("CLAIM", qury_finish)
						MEMBER_INFO_ARRAY(second_qury_sent, each_memb) = qury_finish

						objIncomeSQL = "UPDATE ES.ES_ExParte_IncomeList SET QURY_Sent = '" & qury_finish & "' WHERE [CaseNumber] = '" & MAXIS_case_number & "' and [PersonID] = '" & MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb) & "' and [ClaimNbr] like '" & MEMBER_INFO_ARRAY(unmatched_claim_numb, each_memb) & "%'"

						'Creating objects for Access
						Set objIncomeConnection = CreateObject("ADODB.Connection")
						Set objIncomeRecordSet = CreateObject("ADODB.Recordset")

						'This is the file path for the statistics Access database.
						' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
						objIncomeConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
						objIncomeRecordSet.Open objIncomeSQL, objIncomeConnection

					End If
				Next

				If va_count <> 0 Then
					If last_va_count <> va_count Then
						If last_va_count = 0 Then
							va_excel_created = True
							'Create an Excel file to record members that have VA Income
							Set objVAExcel = CreateObject("Excel.Application")
							objVAExcel.Visible = True
							Set objWorkbook = objVAExcel.Workbooks.Add()
							objVAExcel.DisplayAlerts = True

							'Setting the first 4 col as worker, case number, name, and APPL date
							objVAExcel.Cells(1, 1).Value = "CASE NUMBER"
							objVAExcel.Cells(1, 2).Value = "REF"
							objVAExcel.Cells(1, 3).Value = "NAME"
							objVAExcel.Cells(1, 4).Value = "PMI NUMBER"
							objVAExcel.Cells(1, 5).Value = "SSN"
							objVAExcel.Cells(1, 6).Value = "VA INC TYPE"
							objVAExcel.Cells(1, 7).Value = "VA CLAIM NUMB"
							objVAExcel.Cells(1, 8).Value = "CURR VA INCOME"
							objVAExcel.Cells(1, 9).Value = "Verified VA Income"
							objVAExcel.columns(2).NumberFormat = "@" 		'formatting as text

							FOR i = 1 to 9		'formatting the cells'
								objVAExcel.Cells(1, i).Font.Bold = True		'bold font'
							NEXT

							va_excel_row = 2
							va_inc_count = 0
						End If

						Do
							objVAExcel.Cells(va_excel_row, 1).value = VA_INCOME_ARRAY(va_case_numb_const, va_inc_count)
							objVAExcel.Cells(va_excel_row, 2).value = VA_INCOME_ARRAY(va_ref_numb_const, va_inc_count)
							objVAExcel.Cells(va_excel_row, 3).value = VA_INCOME_ARRAY(va_pers_name_const, va_inc_count)
							objVAExcel.Cells(va_excel_row, 4).value = VA_INCOME_ARRAY(va_pers_pmi_const, va_inc_count)
							objVAExcel.Cells(va_excel_row, 5).value = VA_INCOME_ARRAY(va_pers_ssn_const, va_inc_count)
							objVAExcel.Cells(va_excel_row, 6).value = VA_INCOME_ARRAY(va_inc_type_code_const, va_inc_count) & " - " & VA_INCOME_ARRAY(va_inc_type_info_const, va_inc_count)
							objVAExcel.Cells(va_excel_row, 7).value = VA_INCOME_ARRAY(va_claim_numb_const, va_inc_count)
							objVAExcel.Cells(va_excel_row, 8).value = VA_INCOME_ARRAY(va_prosp_inc_const, va_inc_count)

							va_inc_count = va_inc_count + 1
							va_excel_row = va_excel_row + 1
						Loop until va_inc_count = va_count
					End If
				End If

				If uc_count <> 0 Then
					If last_uc_count <> uc_count Then
						If last_uc_count = 0 Then
							uc_excel_created = True
							'Create an Excel file to record members that have VA Income
							Set objUCExcel = CreateObject("Excel.Application")
							objUCExcel.Visible = True
							Set objWorkbook = objUCExcel.Workbooks.Add()
							objUCExcel.DisplayAlerts = True

							'Setting the first 4 col as worker, case number, name, and APPL date
							objUCExcel.Cells(1, 1).Value = "CASE NUMBER"
							objUCExcel.Cells(1, 2).Value = "REF"
							objUCExcel.Cells(1, 3).Value = "NAME"
							objUCExcel.Cells(1, 4).Value = "PMI NUMBER"
							objUCExcel.Cells(1, 5).Value = "SSN"
							objUCExcel.Cells(1, 6).Value = "UC INC TYPE"
							objUCExcel.Cells(1, 7).Value = "UC CLAIM NUMB"
							objUCExcel.Cells(1, 8).Value = "CURR UC INCOME"
							objUCExcel.Cells(1, 9).Value = "Verified UC Income"
							objUCExcel.columns(2).NumberFormat = "@" 		'formatting as text

							FOR i = 1 to 9		'formatting the cells'
								objUCExcel.Cells(1, i).Font.Bold = True		'bold font'
							NEXT

							uc_excel_row = 2
							uc_inc_count = 0
						End If

						Do
							objUCExcel.Cells(uc_excel_row, 1).value = UC_INCOME_ARRAY(uc_case_numb_const, uc_inc_count)
							objUCExcel.Cells(uc_excel_row, 2).value = UC_INCOME_ARRAY(uc_ref_numb_const, uc_inc_count)
							objUCExcel.Cells(uc_excel_row, 3).value = UC_INCOME_ARRAY(uc_pers_name_const, uc_inc_count)
							objUCExcel.Cells(uc_excel_row, 4).value = UC_INCOME_ARRAY(uc_pers_pmi_const, uc_inc_count)
							objUCExcel.Cells(uc_excel_row, 5).value = UC_INCOME_ARRAY(uc_pers_ssn_const, uc_inc_count)
							objUCExcel.Cells(uc_excel_row, 6).value = UC_INCOME_ARRAY(uc_inc_type_code_const, uc_inc_count) & " - " & UC_INCOME_ARRAY(uc_inc_type_info_const, uc_inc_count)
							objUCExcel.Cells(uc_excel_row, 7).value = UC_INCOME_ARRAY(uc_claim_numb_const, uc_inc_count)
							objUCExcel.Cells(uc_excel_row, 8).value = UC_INCOME_ARRAY(uc_prosp_inc_const, uc_inc_count)

							uc_inc_count = uc_inc_count + 1
							uc_excel_row = uc_excel_row + 1
						Loop until uc_inc_count = uc_count
					End If
				End If

				'save details of the actions into the table
				' objRecordSet.Open "UPDATE"
			End If

			Call back_to_SELF

			prep_status = date
			If appears_ex_parte = False Then
				prep_status = "Not Ex Parte"
				' If mfip_status = "ACTIVE" OR snap_status = "ACTIVE" Then prep_status = "SNAP/MFIP"
			End If

			objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET SelectExParte = '" & appears_ex_parte & "', PREP_Complete = '" & prep_status & "', AllHCisABD = '" & all_hc_is_ABD & "', SSAIncomExist = '" & SSA_income_exists & "', WagesExist = '" & JOBS_income_exists & "', VAIncomeExist = '" & VA_income_exists & "', SelfEmpExists = '" & BUSI_income_exists & "', NoIncome = '" & case_has_no_income & "', EPDonCase = '" & case_has_EPD & "' WHERE CaseNumber = '" & MAXIS_case_number & "'"

			'Creating objects for Access
			Set objUpdateConnection = CreateObject("ADODB.Connection")
			Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

			'This is the file path for the statistics Access database.
			objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
			objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection

		End If
		objRecordSet.MoveNext
	Loop

	For col_to_autofit = 1 to 9
		If va_excel_created = True Then objVAExcel.columns(col_to_autofit).AutoFit()
		If uc_excel_created = True Then objUCExcel.columns(col_to_autofit).AutoFit()
	Next
	' MsgBox "Cases that appear Ex Parte: " & ex_parte_cases_count

    objRecordSet.Close
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing


	end_msg = "BULK Prep Run has been completed."

	'declare the SQL statement that will query the database
	objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"

	'Creating objects for Access
	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'This is the file path for the statistics Access database.
	' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	'Setting a starting value for a list of cases so that every case is bracketed by * on both sides.
	todays_cases_list = "*"
	case_entry = 0      'Setting an incrementor for the array to be filled

	'Open The CASE LIST Table
	'Loop through each item on the CASE LIST Table
	case_count = 0
	ex_parte_count = 0
	Do While NOT objRecordSet.Eof
		case_count = case_count + 1
		If objRecordSet("SelectExParte") = True Then ex_parte_count = ex_parte_count + 1
		objRecordSet.MoveNext
	Loop
    objRecordSet.Close
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing
	percent_ex_parte = ex_parte_count/case_count
	percent_ex_parte = percent_ex_parte * 100
	percent_ex_parte = FormatNumber(percent_ex_parte, 2, -1, 0, -1)

	end_msg = end_msg & vbCr & "Cases appear to have a HC ER scheduled for " & ep_revw_mo & "/" & ep_revw_yr & ": " & case_count
	end_msg = end_msg & vbCr & "Cases that appear to meet Ex Parte Criteria: " & ex_parte_count
	end_msg = end_msg & vbCr & "This appears to be " & percent_ex_parte & "% of cases."
End If

If ex_parte_function = "Phase 1" Then
	' MsgBox 	"In preparation for the HSR completion of a Phase 1 review, the script will complete updates to MAXIS information, to prevent HSRs from having to amnually enter verified information." & vbCr & vbCr &_
	' 		"If the case is potentially Ex Parte, the script will:" & vbCr &_
	' 		" - Read SVES/TPQY" & vbCr &_
	' 		" - Update UNEA and MEDI with SSA information from SVES/TPQY." & vbCr &_
	' 		" - Enter VA Income reported back after verification." & vbCr &_
	' 		" - Create a CASE/NOTE of any information verified and updated in MAXIS." & vbCr &_
	' 		" - Run the case through background." & vbCr &_
	' 		" - Capture details of the income verified and the Eligibility results into the Table to track Ex Parte work." & vbCr & vbCr &_
	' 		"This script will look at each case for the specified review month, preparing the case to be assigned to an HSR for Phase 1 Review of Ex Parte Eligbility." & vbCr  & vbCr &_
	' 		"This script is run on the 1st of the month of the Budget Month."


	MAXIS_footer_month = CM_plus_1_mo
	MAXIS_footer_year = CM_plus_1_yr

	review_date = ep_revw_mo & "/1/" & ep_revw_yr
	review_date = DateAdd("d", 0, review_date)

	va_count = 0
	uc_count = 0
	ex_parte_cases_count = 0

	'Open The CASE LIST Table
	'declare the SQL statement that will query the database
	objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"

	'Creating objects for Access
	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'This is the file path for the statistics Access database.
	' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	'Setting a starting value for a list of cases so that every case is bracketed by * on both sides.
	todays_cases_list = "*"
	case_entry = 0      'Setting an incrementor for the array to be filled

	'Loop through each item on the CASE LIST Table
	Do While NOT objRecordSet.Eof
		If objRecordSet("SelectExParte") = True and IsNull(objRecordSet("Phase1Complete")) = True Then
			'For each case that is indicated as Ex parte, we are going to update the case information
			MAXIS_case_number = objRecordSet("CaseNumber") 		'SET THE MAXIS CASE NUMBER

			Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)

			If case_active = False Then

				objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET Phase1Complete = 'Case not Active' WHERE CaseNumber = '" & MAXIS_case_number & "'"

				'Creating objects for Access
				Set objUpdateConnection = CreateObject("ADODB.Connection")
				Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

				'This is the file path for the statistics Access database.
				objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
				objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection
			Else

				ReDim MEMBER_INFO_ARRAY(memb_last_const, 0)
				Do
					Call navigate_to_MAXIS_screen("STAT", "MEMB")
					EMReadScreen memb_check, 4, 2, 48
				Loop until memb_check = "MEMB"
				Call get_list_of_members

				'Read SVES/TPQY for all persons on a case
				For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
					MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = False
					MEMBER_INFO_ARRAY(tpqy_memb_has_rsdi, each_memb) = False
					MEMBER_INFO_ARRAY(tpqy_rsdi_has_disa, each_memb) = False
					memb_has_railroad = False
					ssi_end_date = ""
					rsdi_end_date = ""

					Call navigate_to_MAXIS_screen("INFC", "SVES")
					EMWriteScreen MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb), 5, 68
					EMWriteScreen "TPQY", 20, 70
					transmit

					EMReadScreen check_TPQY_panel, 4, 2, 53 		'Reads for TPQY panel
					If check_TPQY_panel = "TPQY" Then
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_record, each_memb), 		1, 8, 39
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_record, each_memb), 		1, 8, 65
						EMReadScreen sves_response, 8, 7, 22 		'Return Date
						sves_response = replace(sves_response," ", "/")
						' MsgBox "SSI record - " & MEMBER_INFO_ARRAY(tpqy_ssi_record, each_memb) & vbCr & "RSDI record - " & MEMBER_INFO_ARRAY(tpqy_rsdi_record, each_memb)
					End If
					transmit

					EMReadScreen check_BDXP_panel, 4, 2, 53 		'Reads fro BDXP panel
					If check_BDXP_panel = "BDXP" Then
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb), 	12, 5, 40
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb), 		10, 5, 69
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb), 	2, 6, 19
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb), 	8, 8, 16
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_net_amt, each_memb), 		8, 8, 32
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_railroad_ind, each_memb), 		1, 8, 69
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb), 	5, 11, 69
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb), 	5, 14, 69
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb), 	10, 15, 69
						MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb))
						MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb))
						MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb))
						' MEMBER_INFO_ARRAY(tpqy_rsdi_staus_desc, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_staus_desc, each_memb))
						' MEMBER_INFO_ARRAY(tpqy_rsdi_paydate, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_paydate, each_memb))
						MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb))
						MEMBER_INFO_ARRAY(tpqy_rsdi_net_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_net_amt, each_memb))
						MEMBER_INFO_ARRAY(tpqy_railroad_ind, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_railroad_ind, each_memb))
						MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb))
						MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb))
						MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb) = Trim (MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb))
						MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb), " ", "/01")
						MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb), " ", "/01")
						MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb), " ", "/")
					End If
					transmit

					EMReadScreen check_BDXM_panel, 4, 2, 53 		'Reads for BDXM panel
					If check_BDXM_panel = "BDXM" Then
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_medi_claim_num, each_memb), 			13, 4, 29
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_premium, each_memb), 			7, 6, 64
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb), 				5, 7, 25
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb), 				5, 7, 63
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_buyin_ind, each_memb), 			1, 8, 25
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_buyin_code, each_memb), 			3, 8, 63
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_buyin_start_date, each_memb), 	5, 9, 25
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_a_buyin_stop_date, each_memb), 	5, 9, 63
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_b_premium, each_memb), 			7, 12, 64
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb), 				5, 13, 25
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb), 				5, 13, 63
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_b_buyin_ind, each_memb), 			1, 14, 25
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_Part_b_buyin_code, each_memb), 			3, 14, 63
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_b_buyin_start_date, each_memb), 	5, 15, 25
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_part_b_buyin_stop_date, each_memb), 	5, 15, 63
						MEMBER_INFO_ARRAY(tpqy_medi_claim_num, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_medi_claim_num, each_memb))
						MEMBER_INFO_ARRAY(tpqy_part_a_premium, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_part_a_premium, each_memb))
						MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb) = trim(MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb))
						MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb) = trim(MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb))
						MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb) = trim(MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb))
						MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb) = trim(MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb))
						MEMBER_INFO_ARRAY(tpqy_part_b_premium, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_part_b_premium, each_memb))
						MEMBER_INFO_ARRAY(tpqy_part_a_buyin_code, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_part_a_buyin_code, each_memb))
						MEMBER_INFO_ARRAY(tpqy_Part_b_buyin_code, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_Part_b_buyin_code, each_memb))
						MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb), " ", "/01/")
						MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb), " ", "/01/")
						MEMBER_INFO_ARRAY(tpqy_part_a_buyin_start_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_a_buyin_start_date, each_memb), " ", "/01/")
						MEMBER_INFO_ARRAY(tpqy_part_a_buyin_stop_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_a_buyin_stop_date, each_memb), " ", "/01/")
						MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb), " ", "/01/")
						MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb), " ", "/01/")
						MEMBER_INFO_ARRAY(tpqy_part_b_buyin_start_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_b_buyin_start_date, each_memb), " ", "/01/")
						MEMBER_INFO_ARRAY(tpqy_part_b_buyin_stop_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_part_b_buyin_stop_date, each_memb), " ", "/01/")
					End If
					transmit

					EMReadScreen check_SDXE_panel, 4, 2, 53 		'Reads for SDXE panel
					If check_SDXE_panel = "SDXE" Then
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_claim_numb, each_memb), 		12, 5, 36
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_recip_code, each_memb), 		2, 7, 21
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_recip_desc, each_memb), 		22, 7, 24
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_fed_living, each_memb), 			1, 6, 70
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb), 			3, 8, 21
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_desc, each_memb), 			30, 8, 25
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_cit_ind_code, each_memb), 			1, 7, 70
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_denial_code, each_memb), 		3, 10, 26
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_denial_desc, each_memb), 		40, 10, 30
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_denial_date, each_memb), 		8, 11, 26
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_disa_date, each_memb), 			8, 12, 26
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_SSP_elig_date, each_memb), 		8, 13, 26
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_appeals_code, each_memb), 		1, 11, 65
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_appeals_date, each_memb), 		8, 12, 65
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_appeals_dec_code, each_memb), 	2, 13, 65
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_appeals_dec_date, each_memb), 	8, 14, 65
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_disa_pay_code, each_memb), 		1, 15, 65
						MEMBER_INFO_ARRAY(tpqy_ssi_claim_numb, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_claim_numb, each_memb))
						MEMBER_INFO_ARRAY(tpqy_ssi_recip_desc, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_recip_desc, each_memb))
						MEMBER_INFO_ARRAY(tpqy_ssi_pay_desc, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_desc, each_memb))
						MEMBER_INFO_ARRAY(tpqy_ssi_denial_desc, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_denial_desc, each_memb))
						MEMBER_INFO_ARRAY(tpqy_ssi_denial_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_denial_date, each_memb))
						MEMBER_INFO_ARRAY(tpqy_ssi_disa_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_disa_date, each_memb))
						MEMBER_INFO_ARRAY(tpqy_ssi_SSP_elig_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_SSP_elig_date, each_memb))
						MEMBER_INFO_ARRAY(tpqy_ssi_appeals_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_appeals_date, each_memb))
						MEMBER_INFO_ARRAY(tpqy_ssi_appeals_dec_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_appeals_dec_date, each_memb))
						MEMBER_INFO_ARRAY(tpqy_ssi_denial_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_denial_date, each_memb), " ", "/")
						MEMBER_INFO_ARRAY(tpqy_ssi_disa_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_disa_date, each_memb), " ", "/")
						MEMBER_INFO_ARRAY(tpqy_ssi_SSP_elig_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_SSP_elig_date, each_memb), " ", "/")
						MEMBER_INFO_ARRAY(tpqy_ssi_appeals_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_appeals_date, each_memb), " ", "/")
						MEMBER_INFO_ARRAY(tpqy_ssi_appeals_dec_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_appeals_dec_date, each_memb), " ", "/")
						' MsgBox MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb)
					End If
					transmit

					EMReadScreen check_SDXP_panel, 4, 2, 50 		'Reads for SDXP panel
					If check_SDXP_panel = "SDXP" Then
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_date, each_memb), 			5, 4, 16
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb), 			7, 4, 42
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_over_under_code, each_memb), 	1, 4, 73
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_date, each_memb), 	5, 8, 3
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_amt, each_memb), 	6, 8, 13
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_type, each_memb), 	1, 8, 25
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_date, each_memb), 	5, 9, 3
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_amt, each_memb), 	6, 9, 13
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_type, each_memb), 	1, 9, 25
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_date, each_memb), 	5, 10, 3
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_amt, each_memb), 	6, 10, 13
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_type, each_memb), 	1, 10, 25
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_gross_EI, each_memb), 				8, 5, 66
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_net_EI, each_memb), 				8, 6, 66
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_income_amt, each_memb), 		8, 7, 66
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_pass_exclusion, each_memb), 		8, 8, 66
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_inc_inkind_start, each_memb), 		8, 9, 66
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_inc_inkind_stop, each_memb), 		8, 10, 66
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_rep_payee, each_memb), 				1, 11, 66
						MEMBER_INFO_ARRAY(tpqy_ssi_pay_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_date, each_memb))
						MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb))
						MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_date, each_memb))
						MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_amt, each_memb))
						MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_date, each_memb))
						MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_amt, each_memb))
						MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_date, each_memb))
						MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_amt, each_memb))
						MEMBER_INFO_ARRAY(tpqy_gross_EI, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_gross_EI, each_memb))
						MEMBER_INFO_ARRAY(tpqy_net_EI, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_net_EI, each_memb))
						MEMBER_INFO_ARRAY(tpqy_rsdi_income_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_income_amt, each_memb))
						MEMBER_INFO_ARRAY(tpqy_pass_exclusion, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_pass_exclusion, each_memb))
						MEMBER_INFO_ARRAY(tpqy_inc_inkind_start, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_inc_inkind_start, each_memb))
						MEMBER_INFO_ARRAY(tpqy_inc_inkind_stop, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_inc_inkind_stop, each_memb))
						MEMBER_INFO_ARRAY(tpqy_ssi_pay_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_pay_date, each_memb), " ", "/01/")
						MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_1_date, each_memb), " ", "/")
						MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_2_date, each_memb), " ", "/")
						MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_ssi_pay_hist_3_date, each_memb), " ", "/")
						MEMBER_INFO_ARRAY(tpqy_inc_inkind_start, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_inc_inkind_start, each_memb), " ", "/")
						MEMBER_INFO_ARRAY(tpqy_inc_inkind_stop, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_inc_inkind_stop, each_memb), " ", "/")
					End If
					transmit

					If MEMBER_INFO_ARRAY(tpqy_ssi_record, each_memb) = "Y" Then
						If MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb) = "C01" Then MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = True
						' If MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb) <> "C01" Then MsgBox "STOP"
					End If
					If MEMBER_INFO_ARRAY(tpqy_rsdi_record, each_memb) = "Y" Then
						If MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb) = "C" or MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb) = "E" Then
							MEMBER_INFO_ARRAY(tpqy_memb_has_rsdi, each_memb) = True
							If IsDate(MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb)) = True Then MEMBER_INFO_ARRAY(tpqy_rsdi_has_disa, each_memb) = True
						' If MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb) = "E" Then
						' 	If MEMBER_INFO_ARRAY(tpqy_railroad_ind, each_memb) = "A" Then memb_has_railroad = True
						End If
					End If
					ssi_end_date = ""
					rsdi_end_date = ""

					' objIncomeSQL = "UPDATE ES.ES_ExParte_IncomeList SET TPQY_Response = '" & sves_response & "' WHERE [CaseNumber] = '" & MAXIS_case_number & "' and [PersonID] = '" & MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb) & "' and [ClaimNbr] like '" & MEMBER_INFO_ARRAY(tpqy_ssi_claim_numb, each_memb) & "%'"

					' 'Creating objects for Access
					' Set objIncomeConnection = CreateObject("ADODB.Connection")
					' Set objIncomeRecordSet = CreateObject("ADODB.Recordset")

					' 'This is the file path for the statistics Access database.
					' ' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
					' objIncomeConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
					' objIncomeRecordSet.Open objIncomeSQL, objIncomeConnection

					Call back_to_SELF

				Next

				Do
					Call navigate_to_MAXIS_screen("STAT", "SUMM")
					EMReadScreen summ_check, 4, 2, 46
				Loop until summ_check = "SUMM"
				verif_types = ""

				For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
					' MsgBox "MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) - " & MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) & vbCr &_
					' 		"MEMBER_INFO_ARRAY(tpqy_memb_has_rsdi, each_memb) - " & MEMBER_INFO_ARRAY(tpqy_memb_has_rsdi, each_memb) & vbCr &_
					' 		"MEMBER_INFO_ARRAY(tpqy_rsdi_has_disa, each_memb) - " & MEMBER_INFO_ARRAY(tpqy_rsdi_has_disa, each_memb)

					'Update MAXIS UNEA panels with information from TPQY
					If MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = True Then
						Call find_UNEA_panel(MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), "03", SSI_UNEA_instance, "", SSI_panel_found)

						Call update_unea_pane(SSI_panel_found, "03", MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb), MEMBER_INFO_ARRAY(tpqy_ssi_claim_numb, each_memb) & MEMBER_INFO_ARRAY(tpqy_ssi_recip_code, each_memb), MEMBER_INFO_ARRAY(tpqy_ssi_SSP_elig_date, each_memb), "")
						If InStr(verif_types, "SSI") = 0 Then verif_types = verif_types & "/SSI"
					End If

					If MEMBER_INFO_ARRAY(tpqy_memb_has_rsdi, each_memb) = True Then
						If MEMBER_INFO_ARRAY(tpqy_rsdi_has_disa, each_memb) = True Then
							rsdi_type = "01"
							Call find_UNEA_panel(MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), rsdi_type, RSDI_UNEA_instance, MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb), RSDI_panel_found)
						Else
							rsdi_type = "02"
							Call find_UNEA_panel(MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), rsdi_type, RSDI_UNEA_instance, MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb), RSDI_panel_found)
						End If
						Call update_unea_pane(RSDI_panel_found, rsdi_type, MEMBER_INFO_ARRAY(tpqy_rsdi_net_amt, each_memb), MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb), MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb), "")
						If InStr(verif_types, "RSDI") = 0 Then verif_types = verif_types & "/RSDI"
					End If

					' MsgBox "1"
					'Update MAXIS MEDI panels with information from TPQY
					MEDI_panel_exists = False
					MEMBER_INFO_ARRAY(created_medi, each_memb) = False
					If MEMBER_INFO_ARRAY(tpqy_medi_claim_num, each_memb) <> "" Then
						' MsgBox "1.5"
						EMWriteScreen "MEDI", 20, 71
						transmit
						EMReadScreen medi_check, 4, 2, 44
						Do while medi_check <> "MEDI"
							Call navigate_to_MAXIS_screen("STAT", "MEDI")
							EMReadScreen medi_check, 4, 2, 44
						Loop
						' EMWriteScreen "MEDI", 20, 71
						EMWriteScreen MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), 20, 76
						transmit
						' MsgBox "1.6"
						EMReadScreen total_amt_of_panels, 1, 2, 78			'Checks to make sure there are JOBS panels for this member. If none exists, one will be created
						MEDI_panel_exists = True
						MEDI_active = False
						If total_amt_of_panels = "0" Then MEDI_panel_exists = False
						If MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb) = "" or MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb) = "" Then MEDI_active = True
						part_a_ended = False
						If IsDate(MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb)) = True Then part_a_ended = True
						part_b_ended = False
						If IsDate(MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb)) = True Then part_b_ended = True

						panel_part_a_accurate = False
						panel_part_b_accurate = False
						If MEDI_panel_exists = True Then
							Do
								PF20
								EMReadScreen end_of_list, 9, 24, 14
								' MsgBox end_of_list & " - 1"
							Loop Until end_of_list = "LAST PAGE"
							row = 17
							Do
								EMReadScreen begin_dt_a, 8, row, 24 		'reads part a start date
								begin_dt_a = replace(begin_dt_a, " ", "/")	'reformatting with / for date
								If begin_dt_a = "__/__/__" Then begin_dt_a = "" 		'blank out if not a date

								EMReadScreen end_dt_a, 8, row, 35	'reads part a end date
								end_dt_a =replace(end_dt_a , " ", "/")		'reformatting with / for date
								If end_dt_a = "__/__/__" Then end_dt_a = ""					'blank out if not a date

								If part_a_ended = True Then
									If end_dt_a <> "" Then
										panel_part_a_accurate = True
										Exit Do
									End If
									If end_dt_a = "" and begin_dt_a <> "" Then Exit Do
								Else
									If begin_dt_a <> "" and end_dt_a <> "" Then
										Exit Do
									ElseIf begin_dt_a <> "" and end_dt_a = "" Then
										panel_part_a_accurate = True
										Exit Do
									End If
								End If
								row = row - 1
								' MsgBox "PART A row - " & row
								If row = 14 Then
									PF19
									EMReadScreen begining_of_list, 10, 24, 14
									' MsgBox "begining_of_list - " & begining_of_list & vbcr & "1"
									If begining_of_list = "FIRST PAGE" Then
										Exit Do
									Else
										row = 17
									End If
								End If
							Loop
							Do
								PF19
								EMReadScreen begining_of_list, 10, 24, 14
								' MsgBox "begining_of_list - " & begining_of_list & vbcr & "2"
							Loop Until begining_of_list = "FIRST PAGE"

							Do
								PF20
								EMReadScreen end_of_list, 9, 24, 14
								' MsgBox end_of_list & " - 2"
							Loop Until end_of_list = "LAST PAGE"
							row = 17
							Do
								EMReadScreen begin_dt_b, 8, row, 54 		'reads part a start date
								begin_dt_b = replace(begin_dt_b, " ", "/")	'reformatting with / for date
								If begin_dt_b = "__/__/__" Then begin_dt_b = "" 		'blank out if not a date

								EMReadScreen end_dt_b, 8, row, 65	'reads part a end date
								end_dt_b =replace(end_dt_b , " ", "/")		'reformatting with / for date
								If end_dt_b = "__/__/__" Then end_dt_b = ""					'blank out if not a date

								If part_a_ended = True Then
									If end_dt_b <> "" Then
										panel_part_b_accurate = True
										Exit Do
									End If
									If end_dt_b = "" and begin_dt_b <> "" Then Exit Do
								Else
									If begin_dt_b <> "" and end_dt_b <> "" Then
										Exit Do
									ElseIf begin_dt_b <> "" and end_dt_b = "" Then
										panel_part_b_accurate = True
									End If
								End If
								row = row - 1
								' MsgBox "PART B row - " & row
								If row = 14 Then
									PF19
									EMReadScreen begining_of_list, 10, 24, 14
									' MsgBox "begining_of_list - " & begining_of_list & vbcr & "3"
									If begining_of_list = "FIRST PAGE" Then
										Exit Do
									Else
										row = 17
									End If
								End If
							Loop

						End If
						' MsgBox "2"
						' MsgBox "panel_part_a_accurate - " & panel_part_a_accurate & vbCr & "panel_part_b_accurate - " & panel_part_b_accurate
						If MEDI_panel_exists = True and (panel_part_a_accurate = False or panel_part_b_accurate = False) Then
							If InStr(verif_types, "Medicare") = 0 Then verif_types = verif_types & "/Medicare"
							PF9
							' MsgBox "Edit?"
							If panel_part_a_accurate = False Then
								Do
									PF20
									EMReadScreen end_of_list, 34, 24, 2
								Loop Until end_of_list = "COMPLETE THE PAGE BEFORE SCROLLING"
								row = 17
								Do
									EMReadScreen begin_dt_a, 8, row, 24 		'reads part a start date
									begin_dt_a = replace(begin_dt_a, " ", "/")	'reformatting with / for date
									If begin_dt_a = "__/__/__" Then begin_dt_a = "" 		'blank out if not a date

									EMReadScreen end_dt_a, 8, row, 35	'reads part a end date
									end_dt_a =replace(end_dt_a , " ", "/")		'reformatting with / for date
									If end_dt_a = "__/__/__" Then end_dt_a = ""					'blank out if not a date

									If part_a_ended = True Then
										If end_dt_a <> "" Then Exit Do
										If end_dt_a = "" and begin_dt_a <> "" Then
											MEMBER_INFO_ARRAY(updated_medi_a, each_memb) = True
											EMReadScreen verif_code, 1, row, 47
											If verif_code <> "V" Then
												Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb), row, 35, "YY")
											Else
												If row = 17 Then
													PF20
													Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb), 15, 24, "YY")
													Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb), 15, 35, "YY")
												Else
													Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb), row+1, 24, "YY")
													Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb), row+1, 35, "YY")
												End If
											End If
											Exit Do
										End If
									Else
										If begin_dt_a <> "" and end_dt_a <> "" Then
											MEMBER_INFO_ARRAY(updated_medi_a, each_memb) = True
											If row = 17 Then
												PF20
												Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb), 15, 24, "YY")
											Else
												Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb), row+1, 24, "YY")
											End If
											Exit Do
										ElseIf begin_dt_a <> "" and end_dt_a = "" Then
											Exit Do
										End If
									End If
									row = row - 1
									' MsgBox "PART A row - " & row
									If row = 14 Then
										PF19
										EMReadScreen begining_of_list, 10, 24, 14
										' MsgBox "begining_of_list - " & begining_of_list & vbcr & "1"
										If begining_of_list = "FIRST PAGE" Then
											Exit Do
										Else
											row = 17
										End If
									End If
								Loop
								Do
									PF19
									EMReadScreen begining_of_list, 10, 24, 14
									' MsgBox "begining_of_list - " & begining_of_list & vbcr & "2"
								Loop Until begining_of_list = "FIRST PAGE"
							End If
							If panel_part_b_accurate = False Then
								Do
									PF20
									EMReadScreen end_of_list, 34, 24, 2
								Loop Until end_of_list = "COMPLETE THE PAGE BEFORE SCROLLING"
								row = 17
								Do
									EMReadScreen begin_dt_b, 8, row, 54 		'reads part a start date
									begin_dt_b = replace(begin_dt_b, " ", "/")	'reformatting with / for date
									If begin_dt_b = "__/__/__" Then begin_dt_b = "" 		'blank out if not a date

									EMReadScreen end_dt_b, 8, row, 65	'reads part a end date
									end_dt_b =replace(end_dt_b , " ", "/")		'reformatting with / for date
									If end_dt_b = "__/__/__" Then end_dt_b = ""					'blank out if not a date

									If part_a_ended = True Then
										If end_dt_b <> "" Then Exit Do
										If end_dt_b = "" and begin_dt_b <> "" Then
											MEMBER_INFO_ARRAY(updated_medi_b, each_memb) = True
											EMReadScreen verif_code, 1, row, 47
											If verif_code <> "V" Then
												Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb), row, 65, "YY")
											Else
												If row = 17 Then
													PF20
													Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb), 15, 54, "YY")
													Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb), 15, 65, "YY")
												Else
													Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb), row+1, 54, "YY")
													Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb), row+1, 65, "YY")
												End If
											End If
											Exit Do
										End If
									Else
										If begin_dt_b <> "" and end_dt_b <> "" Then
											MEMBER_INFO_ARRAY(updated_medi_b, each_memb) = True
											If row = 17 Then
												PF20
												Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb), 15, 54, "YY")
											Else
												Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb), row+1, 54, "YY")
											End If
											Exit Do
										ElseIf begin_dt_b <> "" and end_dt_b = "" Then
											Exit Do
										End If
									End If
									row = row - 1
									' MsgBox "PART B row - " & row
									If row = 14 Then
										PF19
										EMReadScreen begining_of_list, 10, 24, 14
										' MsgBox "begining_of_list - " & begining_of_list & vbcr & "3"
										If begining_of_list = "FIRST PAGE" Then
											Exit Do
										Else
											row = 17
										End If
									End If
								Loop
							End If
							transmit
							' MsgBox "STOP AND LOOK AT THE PANEL"
							' PF10
						End If
						' MsgBox "3"

						If MEDI_panel_exists = False and MEMBER_INFO_ARRAY(tpqy_medi_claim_num, each_memb) <> "" Then
							If InStr(verif_types, "Medicare") = 0 Then verif_types = verif_types & "/Medicare"
							If (MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb) <> "" and MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb) = "") or (MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb) <> "" and MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb) = "") Then
								MEMBER_INFO_ARRAY(created_medi, each_memb) = True
								Call write_value_and_transmit("NN", 20, 79)
								medi_claim_array = Null
								medi_claim_array = split(MEMBER_INFO_ARRAY(tpqy_medi_claim_num, each_memb))
								EMWriteScreen medi_claim_array(0), 6, 39
								EMWriteScreen medi_claim_array(1), 6, 43
								EMWriteScreen medi_claim_array(2), 6, 46
								EMWriteScreen left(medi_claim_array(3), 1), 6, 51

								If MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb) <> "" Then
									MEMBER_INFO_ARRAY(updated_medi_a, each_memb) = True
									Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb), 15, 24, "YY")
									If MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb) <> "" Then Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb), 15, 35, "YY")
								End If

								If MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb) <> "" Then
									MEMBER_INFO_ARRAY(updated_medi_b, each_memb) = True
									Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb), 15, 54, "YY")
									If MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb) <> "" Then
										Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb), 15, 65, "YY")
									Else
										If MEMBER_INFO_ARRAY(tpqy_part_b_premium, each_memb) <> "" Then EMWriteScreen MEMBER_INFO_ARRAY(tpqy_part_b_premium, each_memb), 7, 73
										If MEMBER_INFO_ARRAY(tpqy_part_b_premium, each_memb) = "" Then
											If IsDate(MEMBER_INFO_ARRAY(tpqy_part_b_buyin_start_date, each_memb)) = True Then Call create_mainframe_friendly_date(MEMBER_INFO_ARRAY(tpqy_part_b_buyin_start_date, each_memb), 8, 44, "YY")
										End If
									End If
								End If
								transmit
								' MsgBox "STOP AND LOOK AT THE PANEL"
								' PF10
							End If
						End If





					End If
					' MsgBox "4"


					' Show_msg = False
					' If MEDI_panel_exists = True and (panel_part_a_accurate = False or panel_part_b_accurate = False) Then Show_msg = True
					' If MEDI_panel_exists = False and MEMBER_INFO_ARRAY(tpqy_medi_claim_num, each_memb) <> "" Then Show_msg = True
					' If Show_msg = True Then
					' 	MsgBox "MEDI Panel Exists - " & MEDI_panel_exists & vbCr &_
					' 		"TPQY MEDI Claim Number - " & MEMBER_INFO_ARRAY(tpqy_medi_claim_num, each_memb) & vbCr &_
					' 		"MEMBER - " & MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb) & vbCr &_
					' 		"panel_part_a_accurate - " & panel_part_a_accurate & "  -  TPQY Part A Start: " & MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb) & ", Stop: " & MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb) & vbCr &_
					' 		"panel_part_b_accurate - " & panel_part_b_accurate & "  -  TPQY Part B Start: " & MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb) & ", Stop: " & MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb)
					' End If
					'TODO - Update MAXIS UNEA panels with information from the VA Verifications report


				Next

				'Send the case through background
				Call write_value_and_transmit("BGTX", 20, 71)
				EMReadScreen wrap_check, 4, 2, 46
				If wrap_check = "WRAP" Then transmit
				Call back_to_SELF
				' MsgBox "5"

				'CASE/NOTE details of the case information
				If left(verif_types, 1) = "/" Then verif_types = right(verif_types, len(verif_types)-1)
				note_title = "Verification of " & verif_types

				If verif_types <> "" Then
					Call navigate_to_MAXIS_screen("CASE", "NOTE")
					EMReadScreen last_note, 55, 5, 25
					last_note = trim(last_note)

					If last_note <> note_title Then
						start_a_blank_CASE_NOTE
						Call write_variable_in_CASE_NOTE(note_title)
						For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
							If MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = True or MEMBER_INFO_ARRAY(tpqy_memb_has_rsdi, each_memb) = True Then
								Call write_variable_in_CASE_NOTE("Income from SSA for MEMB " & MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb) & " - " & MEMBER_INFO_ARRAY(memb_name_const, each_memb) & ".")
								If MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = True Then
									Call write_variable_in_CASE_NOTE(" * SSI Income of $ " & MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb) & " per month.")
								End If
								If MEMBER_INFO_ARRAY(tpqy_memb_has_rsdi, each_memb) = True Then
									rsdi_inc = "RSDI"
									If MEMBER_INFO_ARRAY(tpqy_rsdi_has_disa, each_memb) = True Then rsdi_inc = "RSDI, Disa"
									Call write_variable_in_CASE_NOTE(" * " & rsdi_inc & " Income of $ " & MEMBER_INFO_ARRAY(tpqy_rsdi_net_amt, each_memb) & " per month.")
								End If
								Call write_variable_in_CASE_NOTE("   - Verified through worker initiated data match.")
								Call write_variable_in_CASE_NOTE("   - UNEA panel updated eff " & MAXIS_footer_month & "/" & MAXIS_footer_year & ".")
							End If
						Next
						For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
							If MEMBER_INFO_ARRAY(updated_medi_a, each_memb) = True or MEMBER_INFO_ARRAY(updated_medi_b, each_memb) = True Then
								Call write_variable_in_CASE_NOTE("Medicare for MEMB " & MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb) & " - " & MEMBER_INFO_ARRAY(memb_name_const, each_memb) & ".")

								If MEMBER_INFO_ARRAY(updated_medi_a, each_memb) = True Then
									If MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb) <> "" Then
										Call write_variable_in_CASE_NOTE("* Medicare Part A ended " & MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb))
									Else
										Call write_variable_in_CASE_NOTE("* Medicare Part A started " & MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb))
									End If
								End If
								If MEMBER_INFO_ARRAY(updated_medi_b, each_memb) = True Then
									If MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb) <> "" Then
										Call write_variable_in_CASE_NOTE("* Medicare Part B ended " & MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb))
									Else
										Call write_variable_in_CASE_NOTE("* Medicare Part B started " & MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb))
										If MEMBER_INFO_ARRAY(tpqy_part_b_premium, each_memb) <> "" Then
											Call write_variable_in_CASE_NOTE("  - Part B Premium: $ " &MEMBER_INFO_ARRAY(tpqy_part_b_premium, each_memb))
										Else
											Call write_variable_in_CASE_NOTE("  - Part B Buy-In Start Date: " & MEMBER_INFO_ARRAY(tpqy_part_b_buyin_start_date, each_memb))
										End If

									End If
								End If
								Call write_variable_in_CASE_NOTE("   - Verified through worker initiated data match.")
								Call write_variable_in_CASE_NOTE("   - MEDI panel updated eff " & MAXIS_footer_month & "/" & MAXIS_footer_year & ".")
							End If
						Next

						call write_variable_in_case_note("---")
						call write_variable_in_case_note(worker_signature)
						call write_variable_in_case_note("Automated Update")

					End If
				End If

				objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET Phase1Complete = '" & date & "' WHERE CaseNumber = '" & MAXIS_case_number & "'"

				'Creating objects for Access
				Set objUpdateConnection = CreateObject("ADODB.Connection")
				Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

				'This is the file path for the statistics Access database.
				objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
				objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection


				' Call back_to_SELF
				' Call MAXIS_background_check

				' 'Read ELIG and MMIS
				' Call navigate_to_MAXIS_screen("ELIG", "HC  ")
				' For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
				' 	elig_row = 8
				' 	Do
				' 		EMReadScreen ref_num, 2, elig_row, 3
				' 		EMReadScreen elig_program, 10, elig_row, 28
				' 		elig_program = trim(elig_program)
				' 		If ref_num = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb) Then
				' 			EMReadScreen result, 8, elig_row, 41
				' 			result = trim(result)


				' 			Exit Do
				' 		End If

				' 	Loop until ref_num = "  " and elig_program = ""
				' Next
				' ' ,[Phase1ELIG_Created]
				' ' ,[Phase1ELIG_Result]
				' ' ,[Phase1ELIG_Prog]
				' ' ,[Phase1ELIG_Type]
				' ' ,[Phase1ELIG_Method]
				' ' ,[Phase1ELIG_IncStandard]
				' ' ,[Phase1ELIG_Waiver]
				' ' ,[Phase1ELIG_SpenddownType]
				' ' ,[Phase1ELIG_SpenddownAmount]


				'Save all details from the income updates and ELIG information into the SQL Table
			End If

		End If
		objRecordSet.MoveNext
	Loop

End If

'Phase 1 Complete BULK Run
'Read the WHOLE Case List Table
'If SelectExParte = 1 And Phase1HSR = NULL - check to see if on REVW and update SelectExParte = 0 and Phase1HSR = DID NOT COMPLETE
'EMail BI when done

If ex_parte_function = "Phase 2" Then
	MsgBox "Phase 2 BULK Run Details to be added later. This functionality will prep cases for HSR Review at Phase 2, which will happen at the beginning of the Processing month (the month before the Review Month)."
End If


'Loop through all the SQL Items and look for the right revew month and year and phase to determine if it's done.

Call script_end_procedure(end_msg)
