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
		MEMBER_INFO_ARRAY(unea_RR_exists, each_memb) = False

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

				If income_type_code = "16" Then
					MEMBER_INFO_ARRAY(unea_RR_exists, each_memb) = True
					ReDim Preserve RR_INCOME_ARRAY(rr_last_const, rr_count)

					RR_INCOME_ARRAY(rr_case_numb_const, rr_count) = MAXIS_case_number
					RR_INCOME_ARRAY(rr_ref_numb_const, rr_count) = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
					RR_INCOME_ARRAY(rr_pers_name_const, rr_count) = MEMBER_INFO_ARRAY(memb_name_const, each_memb)
					RR_INCOME_ARRAY(rr_pers_ssn_const, rr_count) = left(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 3) & "-" & mid(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4, 2) & "-" & right(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4)
					RR_INCOME_ARRAY(rr_pers_pmi_const, rr_count) = MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb)
					RR_INCOME_ARRAY(rr_inc_type_code_const, rr_count) = income_type_code
					RR_INCOME_ARRAY(rr_inc_type_info_const, rr_count) = "Unemployment"
					RR_INCOME_ARRAY(rr_claim_numb_const, rr_count) = claim_num
					EMReadScreen RR_INCOME_ARRAY(rr_prosp_inc_const, rr_count), 8, 13, 68
					RR_INCOME_ARRAY(rr_prosp_inc_const, rr_count) = trim(RR_INCOME_ARRAY(rr_prosp_inc_const, rr_count))
					If RR_INCOME_ARRAY(rr_prosp_inc_const, rr_count) = "________" Then RR_INCOME_ARRAY(rr_prosp_inc_const, rr_count) = "0.00"

					rr_count = rr_count + 1
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


			If panel_type_code = UNEA_type_code and (UNEA_type_code = "01" or UNEA_type_code = "02") Then
				If len(UNEA_claim_number) = len(panel_claim_number) Then
					If UNEA_claim_number = panel_claim_number Then panel_found = True
				ElseIf len(UNEA_claim_number) < len(panel_claim_number) Then
					If UNEA_claim_number = left(panel_claim_number, len(UNEA_claim_number)) Then panel_found = True
				ElseIf len(UNEA_claim_number) > len(panel_claim_number) Then
					If left(UNEA_claim_number, len(panel_claim_number)) = panel_claim_number Then panel_found = True
				End If
			Else
				If panel_type_code = UNEA_type_code Then panel_found = True
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
				EMReadScreen MEMBER_INFO_ARRAY(memb_smi_numb_const, known_membs), 9, 5, 46
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
			EMReadScreen MEMBER_INFO_ARRAY(memb_smi_numb_const, client_count), 9, 5, 46
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
		EMWriteScreen MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb), 7, 38
	End If
	EMWriteScreen MAXIS_case_number, 	11, 38
	EMWriteScreen "Y", 					14, 38
	' MsgBox "check entry"
	transmit  'Now it sends the SVES.

	EMReadScreen duplicate_SVES, 	    7, 24, 2
	If duplicate_SVES = "WARNING" then transmit
	EMReadScreen confirm_SVES, 			6, 24, 2
	' MsgBox "confirm_SVES - " & confirm_SVES
	if confirm_SVES = "RECORD" then
		' PMI_array(SVES_status, item) = True
		qury_finish = date
	Else
		' PMI_array(SVES_status, item) = False
		qury_finish = "FAILED"
	END IF
end function

function update_stat_budg()
	Call navigate_to_MAXIS_screen("STAT", "BUDG")
	EMReadScreen budg_begin_mo, 2, 10, 35
	EMReadScreen budg_begin_yr, 2, 10, 38
	EMReadScreen budg_end_mo, 2, 10, 46
	EMReadScreen budg_end_yr, 2, 10, 49

	If budg_begin_mo <> ep_revw_mo Then
		PF9 		'put the panel in update mode.
		EMWriteScreen ep_revw_mo, 5, 64
		EMWriteScreen ep_revw_yr, 5, 67
		EMWriteScreen ep_end_budg_revw_mo, 5, 72
		EMWriteScreen ep_end_budg_revw_yr, 5, 75
		' reenter_correct_months = False
		transmit

		EMReadScreen edit_message, 56, 24, 2
		edit_message = trim(edit_message)

		If edit_message <> "" Then
			objTextStream.WriteLine "Case: " & MAXIS_case_number & " - BUDG not updated"
			PF10
		End If
		' new_mo_start = budg_begin_mo
		' new_yr_start = budg_begin_yr
		' new_mo_end = budg_end_mo
		' new_yr_end = budg_end_yr
		' end_of_last_budg_date = budg_end_mo & "/1/" & budg_end_yr
		' end_of_last_budg_date = DateAdd("d", 0, end_of_last_budg_date)

		' Do while edit_message <> ""
		' 	' reenter_correct_months = True
		' 	PF9

		' 	next_budg_pd_start = DateAdd("m", 1, end_of_last_budg_date)
		' 	next_budg_pd_end = DateAdd("m", 6, end_of_last_budg_date)
		' 	Call convert_date_into_MAXIS_footer_month(next_budg_pd_start, new_mo_start, new_yr_start)
		' 	Call convert_date_into_MAXIS_footer_month(next_budg_pd_end, new_mo_end, new_yr_end)
		' 	MsgBox "next_budg_pd_start - " & next_budg_pd_start & vbCr & "new_mo_start - " & new_mo_start & vbCr & "new_yr_start - " & new_yr_start & vbcr & vbCr &_
		' 			 "next_budg_pd_end - " & next_budg_pd_end & vbCr & "new_mo_end - " & new_mo_end & vbCr & "new_yr_end - " & new_yr_end & vbcr & vbCr &_
		' 			""


		' 	EMWriteScreen new_mo_start, 5, 64
		' 	EMWriteScreen new_yr_start, 5, 67
		' 	EMWriteScreen new_mo_end, 5, 72
		' 	EMWriteScreen new_yr_end, 5, 75
		' 	MsgBox "Loop look"
		' 	transmit

		' 	Call back_to_SELF
		' 	Call MAXIS_background_check
		' 	Call navigate_to_MAXIS_screen("STAT", "BUDG")
		' 	PF9
		' 	EMWriteScreen ep_revw_mo, 5, 64
		' 	EMWriteScreen ep_revw_yr, 5, 67
		' 	EMWriteScreen ep_end_budg_revw_mo, 5, 72
		' 	EMWriteScreen ep_end_budg_revw_yr, 5, 75
		' 	transmit

		' 	EMReadScreen edit_message, 56, 24, 2
		' 	edit_message = trim(edit_message)
		' 	end_of_last_budg_date = next_budg_pd_end
		' Loop

		' If reenter_correct_months = True Then
		' 	PF9
		' 	EMWriteScreen ep_revw_mo, 5, 64
		' 	EMWriteScreen ep_revw_yr, 5, 67
		' 	EMWriteScreen ep_end_budg_revw_mo, 5, 72
		' 	EMWriteScreen ep_end_budg_revw_yr, 5, 75
		' 	transmit
		' End If
	End If
	' transmit
	' MsgBox "look"
end function


function update_unea_pane(panel_found, unea_type, income_amount, claim_number, start_date, end_date, last_pay)
	panel_in_edit_mode = False
	If panel_found = False and end_date = "" Then
		Call write_value_and_transmit("NN", 20, 79)
		panel_in_edit_mode = True
	ElseIf panel_found = True Then
		PF9
		panel_in_edit_mode = True
	End If
	If panel_in_edit_mode = True Then
		If claim_number <> "" Then
			Call clear_line_of_text(6, 37)
			EMWriteScreen claim_number, 6, 37
		End If
		EMWriteScreen unea_type, 5, 37
		EMWriteScreen "7", 5, 65		'Write Verification Worker Initiated Verfication "7"
		If unea_type = "11" or unea_type = "12" or unea_type = "13" or unea_type = "38" Then EMWriteScreen "6", 5, 65
		If panel_found = False Then
			Call create_mainframe_friendly_date(start_date, 7, 37, "YY") 	'income start date (SSI: ssi_SSP_elig_date, RSDI: intl_entl_date)
		Else
			EMReadScreen start_date, 8, 7, 37
			start_date = replace(start_date, " ", "/")
			start_date = DateAdd("d", 0 , start_date)
		End If

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
		retro_date = CM_minus_1_mo & "/1/" & CM_minus_1_yr
		retro_date = DateAdd("d", 0, retro_date)

		If end_date <> "" Then
			Call create_mainframe_friendly_date(end_date, 7, 68, "YY")	'income end date (SSI: ssi_denial_date, RSDI: susp_term_date)
			Call convert_date_into_MAXIS_footer_month(end_date, footer_month_end, footer_year_end)

			If footer_month_end = CM_plus_1_mo and footer_year_end = CM_plus_1_yr Then
				Call create_mainframe_friendly_date(end_date, 13, 54, "YY")
				EMWriteScreen last_pay, 13, 68
			End If
		Else
			' MsgBox "retro_date - " & retro_date & vbCr & "start_date - " & start_date & vbCr & "DateDiff - " & DateDiff("d", retro_date, start_date)
			'TODO - this retro date thing failed
			If DateDiff("d", start_date, retro_date) >= 0 Then
				EMWriteScreen CM_minus_1_mo, 13, 25 'hardcoded dates
				EMWriteScreen "01", 13, 28
				EMWriteScreen CM_minus_1_yr, 13, 31 'hardcoded dates
				EMWriteScreen income_amount, 13, 39		'TODO: Testing values
			End If
			' MsgBox "STOP and look"
			EMWriteScreen CM_plus_1_mo, 13, 54 'hardcoded dates
			EMWriteScreen "01", 13, 57
			EMWriteScreen CM_plus_1_yr, 13, 60 'hardcoded dates
			EMWriteScreen income_amount, 13, 68		'TODO: Testing values (income_amt which = rsdi_gross_amt or ssi_gross_amt )

			Call write_value_and_transmit("X", 6, 56)
			Call clear_line_of_text(9, 65)
			EMWriteScreen income_amount, 9, 65		'TODO: Testing values (rsdi_gross_amt or ssi_gross_amt )
			EMWriteScreen "1", 10, 63		'code for pay frequency
			Do
				transmit
				EMReadScreen HC_popup, 9, 7, 41
				' If HC_popup = "HC Income" then transmit
			Loop until HC_popup <> "HC Income"
		End If

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
Const memb_smi_numb_const	= 12

Const unea_type_01_esists	= 20
Const unea_type_02_esists	= 21
Const unea_type_03_esists	= 22
Const unea_type_16_esists	= 23
Const unmatched_claim_numb	= 24
Const unea_VA_exists		= 25
Const unea_UC_exists		= 26
Const unea_RR_exists		= 27

Const sves_qury_sent		= 35
Const second_qury_sent		= 36
Const sves_tpqy_response	= 37
Const sql_uc_income_exists	= 38
Const sql_va_income_exists	= 39
Const sql_rr_income_exists	= 40
Const tpqy_date_of_death	= 41

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
Const tpqy_ssi_last_pay_date		= 107
Const tpqy_ssi_is_ongoing			= 108
COnst tpqy_ssi_last_pay_amt			= 109

Const tpqy_memb_has_ssi				= 110
Const tpqy_memb_has_rsdi			= 111
Const tpqy_rsdi_has_disa			= 112
Const created_medi					= 113
Const updated_medi_a				= 114
Const updated_medi_b				= 115


Const memb_last_const 		= 120

Dim MEMBER_INFO_ARRAY()


Const va_case_numb_const 		= 0
Const va_ref_numb_const 		= 1
Const va_pers_name_const		= 2
Const va_pers_pmi_const			= 3
Const va_pers_ssn_const			= 4
Const va_inc_type_code_const 	= 5
Const va_inc_type_info_const	= 6
Const va_claim_numb_const 		= 7
Const va_prosp_inc_const 		= 8
Const va_end_date_const			= 9
Const va_panel_updated_const 	= 10
Const va_last_const 			= 11

Dim VA_INCOME_ARRAY()
ReDim VA_INCOME_ARRAY(va_last_const, 0)

Const uc_case_numb_const 		= 0
Const uc_ref_numb_const 		= 1
Const uc_pers_name_const		= 2
Const uc_pers_pmi_const			= 3
Const uc_pers_ssn_const			= 4
Const uc_inc_type_code_const 	= 5
Const uc_inc_type_info_const	= 6
Const uc_claim_numb_const 		= 7
Const uc_prosp_inc_const 		= 8
Const uc_end_date_const			= 9
Const uc_panel_updated_const 	= 10
Const uc_last_const 			= 11

Dim UC_INCOME_ARRAY()
ReDim UC_INCOME_ARRAY(uc_last_const, 0)

Const rr_case_numb_const 		= 0
Const rr_ref_numb_const 		= 1
Const rr_pers_name_const		= 2
Const rr_pers_pmi_const			= 3
Const rr_pers_ssn_const			= 4
Const rr_inc_type_code_const 	= 5
Const rr_inc_type_info_const	= 6
Const rr_claim_numb_const 		= 7
Const rr_prosp_inc_const 		= 8
Const rr_end_date_const			= 9
Const rr_panel_updated_const 	= 10
Const rr_last_const 			= 11

Dim RR_INCOME_ARRAY()
ReDim RR_INCOME_ARRAY(rr_last_const, 0)


'Setting constants
Const adOpenStatic = 3
Const adLockOptimistic = 3

'These are the constants that we need to create tables in Excel
Const xlSrcRange = 1
Const xlYes = 1

'END DECLARATIONS BLOCK ====================================================================================================


'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone

Confirm_Process_to_Run_btn	= 200
incorrect_process_btn		= 100
end_msg = "DONE"

MAXIS_footer_month = CM_plus_1_mo		'We are always operating in Current Month plus 1 while runing this script
MAXIS_footer_year = CM_plus_1_yr

ex_parte_folder = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\Ex Parte"


If Day(date) < 1 Then ex_parte_function = "Prep"

'DISPLAYS DIALOG

DO
	DO
		DO
			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 401, 255, "Ex Parte Report"
				DropListBox 300, 25, 90, 15, "Select one..."+chr(9)+"Prep 1"+chr(9)+"Prep 2"+chr(9)+"Phase 1"+chr(9)+"Phase 2"+chr(9)+"ADMIN Review"+chr(9)+"FIX LIST"+chr(9)+"Check REVW information on Phase 1 Cases"+chr(9)+"DHS Data Validation", ex_parte_function
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

		If ex_parte_function <> "ADMIN Review" Then
			allow_bulk_run_use = False
			If user_ID_for_validation = "CALO001" Then allow_bulk_run_use = True
			If user_ID_for_validation = "ILFE001" Then allow_bulk_run_use = True
			If user_ID_for_validation = "MARI001" Then allow_bulk_run_use = True
			If user_ID_for_validation = "MEGE001" Then allow_bulk_run_use = True

			If allow_bulk_run_use = False Then script_end_procedure("Ex Parte Report functionality for completing Ex Parte actions and list review is locked. The script will now end.")

			If ex_parte_function = "Prep 1" or ex_parte_function = "Prep 2" or ex_parte_function = "FIX LIST" or ex_parte_function = "DHS Data Validation" Then
				ep_revw_mo = right("00" & DatePart("m",	DateAdd("m", 3, date)), 2)
				ep_revw_yr = right(DatePart("yyyy",	DateAdd("m", 3, date)), 2)
				' ep_revw_mo = "07"
				' ep_revw_yr = "23"

			End If
			If ex_parte_function = "Phase 1" or ex_parte_function = "Check REVW information on Phase 1 Cases" Then
				ep_revw_mo = right("00" & DatePart("m",	DateAdd("m", 2, date)), 2)
				ep_revw_yr = right(DatePart("yyyy",	DateAdd("m", 2, date)), 2)
			End If
			If ex_parte_function = "Phase 2" Then
				ep_revw_mo = right("00" & DatePart("m",	DateAdd("m", 1, date)), 2)
				ep_revw_yr = right(DatePart("yyyy",	DateAdd("m", 1, date)), 2)
				ep_end_budg_revw_mo = right("00" & DatePart("m",	DateAdd("m", 6, date)), 2)
				ep_end_budg_revw_yr = right(DatePart("yyyy",	DateAdd("m", 6, date)), 2)
			End If

			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 341, 165, "Confirm Ex Parte process"
				EditBox 600, 700, 10, 10, fake_edit_box
				Checkbox 10, 115, 330, 10, "Check here to clear any previous 'In Progress' statuses on cases in the Data Table.", reset_in_Progress
				ButtonGroup ButtonPressed
					PushButton 10, 145, 210, 15, "CONFIRMED! This is the correct Process and Review Month", Confirm_Process_to_Run_btn
					PushButton 230, 145, 100, 15, "Incorrect Process/Month", incorrect_process_btn
				Text 10, 10, 225, 10, "You are running the Ex Parte Function " & ex_parte_function
				Text 10, 25, 190, 10, "This will run for the Ex Parte Review month of " & ep_revw_mo & "/" & ep_revw_yr

				If ex_parte_function = "Prep 1" Then
					GroupBox 5, 40, 240, 60, "Tasks to be Completed:"
					Text 20, 55, 190, 10, "Collect any Case Criteria not available in Info Store."
					Text 20, 65, 175, 10, "Send SVES/QURY for all members on all cases."
					Text 20, 75, 200, 10, "Generate a UC, VA, and RR Verif Report for OS Staff completion."
					Text 20, 85, 200, 10, "Create a list of SMRT ending members."
				End If
				If ex_parte_function = "DHS Data Validation" Then
					GroupBox 5, 40, 240, 60, "Tasks to be Completed:"
					Text 20, 55, 190, 10, "Compare Hennepin Ex Parte list to the cases from the DHS list."
				End If
				If ex_parte_function = "Prep 2" Then
					GroupBox 5, 40, 240, 50, "Tasks to be Completed:"
					Text 20, 55, 270, 10, "Read SVES/TPQY and update UNEA with the response information."
					Text 20, 65, 270, 10, "Send SVES/QURY for all members whose TPQY indicates a second associated RSDI Claim."
					Text 20, 75, 270, 10, "Generate a list of all members with a Date of Death in TPQY."
				End If
				If ex_parte_function = "FIX LIST" Then
					Text 20, 55, 270, 10, "FIX LIST HAS NO SET FUNCTIONALITY"
				End If
				If ex_parte_function = "Phase 1" Then
					' Text 210, 10, 75, 10, "Date of Prep 2 Run:"
					' EditBox 280, 5, 50, 15, prep_phase_2_run_date
					GroupBox 5, 40, 295, 70, "Tasks to be Completed:"
					Text 20, 55, 245, 10, "Read SVES/TPQY Response, Update STAT with detail, enter CASE/NOTE."
					Text 20, 65, 270, 10, "Udate STAT with UC or VA Verifications provided from OS Report and CASE/NOTE."
					Text 20, 75, 125, 10, "Run each case through Background."
					Text 20, 85, 200, 10, "Read and Record in the SQL Table the ELIG information."
					Text 20, 95, 225, 10, "Read and Record in the SQL Table the detail of MMIS Open Spans."
				End If
				If ex_parte_function = "Check REVW information on Phase 1 Cases" Then
					GroupBox 5, 40, 240, 50, "Tasks to be Completed:"
					Text 20, 55, 270, 10, "Pull a list of all cases for the Ex Parte month from the SQL Table."
					Text 20, 65, 270, 10, "Pull all cases from REPT/REVS for the Ex Parte month from MAXIS."
					Text 20, 75, 270, 10, "Read Ex Parte information from STAT/REVW."
					Text 20, 85, 270, 10, "COnnect cases on both lists and output everything to Excel"
				End If
				If ex_parte_function = "Phase 2" Then
					GroupBox 5, 40, 305, 60, "Tasks to be Completed:"
					Text 20, 55, 285, 10, "Check DAIL, CASE/NOTE, STAT for any updates since Phase 1 Ex Parte Determination."
					Text 20, 65, 145, 10, "Record in SQL Table any Updates found."
					Text 20, 75, 125, 10, "Run each case through Background."
					Text 20, 85, 200, 10, "Read and Record in the SQL Table the ELIG information."
				End If

				Text 10, 130, 330, 10, "Review the process datails and ex parte review month to confirm this is the correct run to complete."
			EndDialog

			Dialog Dialog1
			cancel_without_confirmation
			' If IsDate(prep_phase_2_run_date) = False and ex_parte_function = "Phase 1" then
			' 	ButtonPressed = "Loop"
			' 	MsgBox "You must enter a date for the Prep 2 run"
			' End If

			If ButtonPressed = OK Then ButtonPressed = Confirm_Process_to_Run_btn
		Else
			ButtonPressed = Confirm_Process_to_Run_btn
		End If

	Loop until ButtonPressed = Confirm_Process_to_Run_btn
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in


If ex_parte_function = "ADMIN Review" Then
	'this functionality is meant to review the status of cases on the SQL data list. This can help track the progress on Ex Parte cases.

	'This functionality is locked down and only available for use by certain staff.
	allow_admin_use = False
	If user_ID_for_validation = "CALO001" Then allow_admin_use = True
	If user_ID_for_validation = "ILFE001" Then allow_admin_use = True
	If user_ID_for_validation = "MARI001" Then allow_admin_use = True
	If user_ID_for_validation = "MEGE001" Then allow_admin_use = True
	If user_ID_for_validation = "LALA004" Then allow_admin_use = True
	If user_ID_for_validation = "WFX901" Then allow_admin_use = True
	If user_ID_for_validation = "BETE001" Then allow_admin_use = True

	If allow_admin_use = False Then script_end_procedure("ADMIN function for reviewing Ex Parte Functionality is locked. The script will now end.")

	'First we need to set the dates for each phase of Ex Parte.
	current_month_revw = CM_mo & "/1/" & CM_yr
	next_month_revw = CM_plus_1_mo & "/1/" & CM_plus_1_yr
	month_after_next_revw = CM_plus_2_mo & "/1/" & CM_plus_2_yr
	phase_one_hard_stop_date = CM_mo & "/15/" & CM_yr
	prep_month_revw = CM_plus_3_mo & "/1/" & CM_plus_3_yr
	current_month_revw = DateAdd("d", 0, current_month_revw)
	next_month_revw = DateAdd("d", 0, next_month_revw)
	month_after_next_revw = DateAdd("d", 0, month_after_next_revw)
	phase_one_hard_stop_date = DateAdd("d", 0, phase_one_hard_stop_date)
	prep_month_revw = DateAdd("d", 0, prep_month_revw)

	PREP_PHASE_MO = CM_plus_3_mo & "/" & CM_plus_3_yr		'now we are creating strings with the months to display which month is in which phase during the dialog
	PHASE_ONE_MO = CM_plus_2_mo & "/" & CM_plus_2_yr
	PHASE_TWO_MO = CM_plus_1_mo & "/" & CM_plus_1_yr
	COMPLETED_MO = CM_mo & "/" & CM_yr

	Phase_one_hard_stop_passed = False						'defining if we have passed er cut off for Phase 1 work
	If DateDiff("d", phase_one_hard_stop_date, date) > 0 Then Phase_one_hard_stop_passed = True

	'declare the SQL statement that will query the database - we need to pull cases from 3 different review months.
	objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE HCEligReviewDate = '" & next_month_revw & "' or HCEligReviewDate = '" & month_after_next_revw & "' or HCEligReviewDate = '" & prep_month_revw & "'"

	'Creating objects for Access
	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'Opening the SQL data path for Ex Parte
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	'This array is to maintain a list of workers that are processing Ex Parte
	const worker_numb_const 		= 0
	const worker_name_const 		= 1
	const case_complete_p1_count 	= 2
	const case_complete_p2_count	= 3
	const case_phase_const 			= 4
	Dim HSR_WORK_ARRAY()
	ReDim HSR_WORK_ARRAY(case_phase_const, 0)

	'Setting intial numbers as counts for Ex parte case scenarios and types.
	next_month_er_count = 0
	next_month_still_expt = 0
	next_month_hsr_phase2_complete_count = 0
	next_month_need_to_work = 0
	next_month_need_more_review = 0
	next_month_app_count = 0
	next_month_rescheduled_count = 0
	next_month_closed_xfer_count = 0

	month_after_next_er_count = 0
	month_after_next_expt_at_prep = 0
	month_after_next_hsr_phase1_complete_count = 0
	month_after_next_complete_and_expt = 0
	month_after_next_still_expt = 0
	month_after_next_need_to_work = 0

	prep_month_er_count = 0
	prep_month_still_need_eval = 0
	prep_month_expt_at_prep = 0

	'Starting values for looking through all of the cases.
	list_of_hsrs = " "
	hsr_count = 0
	Do While NOT objRecordSet.Eof		'here is where we look at all of the cases to count where everything is at and determine the workers
		'PHASE 2 CASE INFORMATION
		If DateDiff("d", objRecordSet("HCEligReviewDate"), next_month_revw) = 0 Then
			next_month_er_count = next_month_er_count + 1
			If objRecordSet("SelectExParte") = True Then next_month_still_expt = next_month_still_expt + 1
			If IsNull(objRecordSet("Phase2HSR")) = False and trim(objRecordSet("Phase2HSR")) <> "" Then
				next_month_hsr_phase2_complete_count = next_month_hsr_phase2_complete_count + 1
				If objRecordSet("ExParteAfterPhase2") = "REVIEW" Then next_month_need_more_review = next_month_need_more_review + 1
				If objRecordSet("ExParteAfterPhase2") = "Approved as Ex Parte" Then next_month_app_count = next_month_app_count + 1
				If objRecordSet("ExParteAfterPhase2") = "Closed HC" Then next_month_closed_xfer_count = next_month_closed_xfer_count + 1
				If objRecordSet("ExParteAfterPhase2") = "Case not in 27" Then next_month_closed_xfer_count = next_month_closed_xfer_count + 1
				If InStr(objRecordSet("ExParteAfterPhase2"), "ER Scheduled") <> 0 Then next_month_rescheduled_count = next_month_rescheduled_count + 1

				case_phase_two_hsr = objRecordSet("Phase2HSR")
				If InStr(list_of_hsrs, case_phase_two_hsr) = 0 Then
					ReDim Preserve HSR_WORK_ARRAY(case_phase_const, hsr_count)
					HSR_WORK_ARRAY(worker_numb_const, hsr_count) = case_phase_two_hsr
					HSR_WORK_ARRAY(case_complete_p1_count, hsr_count) = 0
					HSR_WORK_ARRAY(case_complete_p2_count, hsr_count) = 1
					hsr_count = hsr_count + 1
					list_of_hsrs = list_of_hsrs & case_phase_two_hsr & " "
				Else
					For each_worker = 0 to UBound(HSR_WORK_ARRAY, 2)
						If HSR_WORK_ARRAY(worker_numb_const, each_worker) = case_phase_two_hsr Then HSR_WORK_ARRAY(case_complete_p2_count, each_worker) = HSR_WORK_ARRAY(case_complete_p2_count, each_worker) + 1
					Next
				End If
			Else
				If objRecordSet("SelectExParte") = True Then next_month_need_to_work = next_month_need_to_work + 1
			End If
			' If objRecordSet("SelectExParte") = True Then month_after_next_still_expt = month_after_next_still_expt + 1
		End If

		'PHASE 1 CASE INFORMATION
		If DateDiff("d", objRecordSet("HCEligReviewDate"), month_after_next_revw) = 0 Then
			month_after_next_er_count = month_after_next_er_count + 1
			If IsDate(objRecordSet("PREP_Complete")) = True and IsDate(objRecordSet("Phase1Complete")) = True Then month_after_next_expt_at_prep = month_after_next_expt_at_prep + 1
			If IsNull(objRecordSet("Phase1HSR")) = False and trim(objRecordSet("Phase1HSR")) <> "" Then
				month_after_next_hsr_phase1_complete_count = month_after_next_hsr_phase1_complete_count + 1
				If objRecordSet("SelectExParte") = True Then month_after_next_complete_and_expt = month_after_next_complete_and_expt + 1
				case_phase_one_hsr = objRecordSet("Phase1HSR")
				If InStr(list_of_hsrs, case_phase_one_hsr) = 0 Then
					ReDim Preserve HSR_WORK_ARRAY(case_phase_const, hsr_count)
					HSR_WORK_ARRAY(worker_numb_const, hsr_count) = case_phase_one_hsr
					HSR_WORK_ARRAY(case_complete_p1_count, hsr_count) = 1
					HSR_WORK_ARRAY(case_complete_p2_count, hsr_count) = 0
					hsr_count = hsr_count + 1
					list_of_hsrs = list_of_hsrs & case_phase_one_hsr & " "
				Else
					For each_worker = 0 to UBound(HSR_WORK_ARRAY, 2)
						If HSR_WORK_ARRAY(worker_numb_const, each_worker) = case_phase_one_hsr Then HSR_WORK_ARRAY(case_complete_p1_count, each_worker) = HSR_WORK_ARRAY(case_complete_p1_count, each_worker) + 1
					Next
				End If
			Else
				If objRecordSet("SelectExParte") = True Then month_after_next_need_to_work = month_after_next_need_to_work + 1
			End If
			If objRecordSet("SelectExParte") = True Then month_after_next_still_expt = month_after_next_still_expt + 1
		End If

		'PREP CASE INFORMATION
		If DateDiff("d", objRecordSet("HCEligReviewDate"), prep_month_revw) = 0 Then
			prep_month_er_count = prep_month_er_count + 1
			If objRecordSet("SelectExParte") = True Then prep_month_expt_at_prep = prep_month_expt_at_prep + 1
			If IsNull(objRecordSet("PREP_Complete")) = True or objRecordSet("PREP_Complete") = "" Then prep_month_still_need_eval = prep_month_still_need_eval + 1
		End If

		objRecordSet.MoveNext		'go to the next case
	Loop
    objRecordSet.Close				'close the data connection
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing

	'here we calculate the percentage of cases for a number of the counts
	function calculate_percent(numerator, denominator, percent)
		percent = numerator/denominator
		percent = percent * 100
		percent = FormatNumber(percent, 2, -1, 0, -1)
	end function
	If prep_month_er_count <> 0 Then call calculate_percent(prep_month_expt_at_prep, prep_month_er_count, prep_month_percent_ex_parte_pcnt)
	call calculate_percent(month_after_next_expt_at_prep, month_after_next_er_count, month_after_next_initially_expt_pcnt)
	call calculate_percent(month_after_next_hsr_phase1_complete_count, month_after_next_expt_at_prep, month_after_next_processed_pcnt)
	call calculate_percent(month_after_next_need_to_work, month_after_next_expt_at_prep, month_after_next_waiting_pcnt)
	If month_after_next_hsr_phase1_complete_count <> 0 Then
		call calculate_percent(month_after_next_complete_and_expt, month_after_next_hsr_phase1_complete_count, month_after_next_complete_and_expt_pcnt)
	End If

	call calculate_percent(next_month_still_expt, next_month_er_count, next_month_initially_expt_pcnt)
	call calculate_percent(next_month_hsr_phase2_complete_count, next_month_still_expt, next_month_processed_pcnt)
	call calculate_percent(next_month_need_to_work, next_month_still_expt, next_month_waiting_pcnt)

	If next_month_hsr_phase2_complete_count <> 0 Then
		call calculate_percent(next_month_app_count, next_month_hsr_phase2_complete_count, next_month_app_pcnt)
		call calculate_percent(next_month_rescheduled_count, next_month_hsr_phase2_complete_count, next_month_rescheduled_pcnt)
		call calculate_percent(next_month_closed_xfer_count, next_month_hsr_phase2_complete_count, next_month_closed_xfer_pcnt)
		call calculate_percent(next_month_need_more_review, next_month_hsr_phase2_complete_count, next_problem_pcnt)
	End If

	'Now we are going to put a name to the worker ID that is entered int he Ex parte list
	SQL_table = "SELECT * from ES.V_ESAllStaff"				'identifying the table that stores the ES Staff user information

	'Creating objects for Access
	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'This is the file path the data tables
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open SQL_table, objConnection							'Here we connect to the data tables

	Do While NOT objRecordSet.Eof										'now we will loop through each item listed in the table of ES Staff
		Name_array = ""
		For each_worker = 0 to UBound(HSR_WORK_ARRAY, 2)
			If HSR_WORK_ARRAY(worker_numb_const, each_worker) = objRecordSet("EmpLogOnID") Then		'If the ID number is found, we will get the name
				HSR_WORK_ARRAY(worker_name_const, each_worker) = objRecordSet("EmpFullName")
				If InStr(HSR_WORK_ARRAY(worker_name_const, each_worker), ",") <> 0 Then				'this will format the name to be easier to read in the dialog display
					Name_array = split(HSR_WORK_ARRAY(worker_name_const, each_worker), ",")
					HSR_WORK_ARRAY(worker_name_const, each_worker) = trim(Name_array(1)) & " " & trim(Name_array(0))
				End If
			End If
			If HSR_WORK_ARRAY(worker_numb_const, each_worker) = "BULK Script" Then HSR_WORK_ARRAY(worker_name_const, each_worker) = "BULK Script"	'this is for the cases that were updated by a BULK run
		Next
		objRecordSet.MoveNext											'Going to the next row in the table
	Loop

	'Now we disconnect from the table and close the connections
	objRecordSet.Close
	objConnection.Close
	Set objRecordSet=nothing
	Set objConnection=nothing

	'Now we need to resize the dialog
	phase_1_factor = 0
	phase_2_factor = 0
	For each_worker = 0 to UBound(HSR_WORK_ARRAY, 2)
		If HSR_WORK_ARRAY(case_complete_p1_count, each_worker) <> 0 Then phase_1_factor = phase_1_factor + 1
		If HSR_WORK_ARRAY(case_complete_p2_count, each_worker) <> 0 Then phase_2_factor = phase_2_factor + 1
	Next
	If phase_1_factor < 6 Then phase_1_factor = 6
	If phase_2_factor < 6 Then phase_2_factor = 6
	If phase_1_factor mod 2 = 1 Then phase_1_factor = phase_1_factor + 1
	If phase_2_factor mod 2 = 1 Then phase_2_factor = phase_2_factor + 1
	dlg_len = 180 + (phase_1_factor/2)*10 + (phase_2_factor/2)*10

	If dlg_len < 180 Then dlg_len = 185

	'display the counts and information gathered from the data list in a dialog
	BeginDialog Dialog1, 0, 0, 500, dlg_len, "Ex Parte Work Details"
		'PREP PHASE
		GroupBox 5, 10, 250, 50, "PREP Phase - " & PREP_PHASE_MO
		Text 15, 25, 155, 10, "Total Cases with HC ER in " & PREP_PHASE_MO & ": " & prep_month_er_count
		If prep_month_expt_at_prep = 0 Then Text 15, 40, 155, 10, "PREP Run not completed."
		If prep_month_expt_at_prep <> 0 Then
			Text 15, 35, 155, 10, "Case that appear Ex Parte Eligible: " & prep_month_expt_at_prep
			Text 175, 35, 75, 10, "Percent: " & prep_month_percent_ex_parte_pcnt & " %"
			If prep_month_still_need_eval <> 0 Then
				Text 25, 45, 155, 10, "Cases that still need PREP run: " & prep_month_still_need_eval
			End If
		End If

		'PHASE 1
		Text 15, 85, 155, 10, "Total Cases with HC ER in " & PHASE_ONE_MO & ": " & month_after_next_er_count
		Text 15, 100, 175, 10, "Cases that appeared Ex Parte at PREP: " & month_after_next_expt_at_prep		'" - XX%"
		Text 115, 110, 175, 10, "Percent: " & month_after_next_initially_expt_pcnt & " %"
		Text 15, 125, 165, 10, "Cases with Phase 1 completed by HSR: " & month_after_next_hsr_phase1_complete_count
		Text 115, 135, 165, 10, "Percent: " & month_after_next_processed_pcnt & " %"
		Text 15, 150, 205, 10, "Cases processed and passed: " & month_after_next_complete_and_expt & "    ( " & month_after_next_complete_and_expt_pcnt & " % )"
		' Text 145, 150, 75, 10, "( " & month_after_next_complete_and_expt_pcnt & " % )"
		If Phase_one_hard_stop_passed = True Then Text 15, 165, 165, 10, "Phase One Processing has stopped."

		y_pos = 85
		x_pos = 185
		For each_worker = 0 to UBound(HSR_WORK_ARRAY, 2)
			If HSR_WORK_ARRAY(case_complete_p1_count, each_worker) <> 0 Then
				If len(HSR_WORK_ARRAY(case_complete_p1_count, each_worker)) = 4 Then Text x_pos, y_pos, 15, 10, HSR_WORK_ARRAY(case_complete_p1_count, each_worker)
				If len(HSR_WORK_ARRAY(case_complete_p1_count, each_worker)) = 3 Then Text x_pos+5, y_pos, 15, 10, HSR_WORK_ARRAY(case_complete_p1_count, each_worker)
				If len(HSR_WORK_ARRAY(case_complete_p1_count, each_worker)) = 2 Then Text x_pos+10, y_pos, 15, 10, HSR_WORK_ARRAY(case_complete_p1_count, each_worker)
				If len(HSR_WORK_ARRAY(case_complete_p1_count, each_worker)) = 1 Then Text x_pos+15, y_pos, 15, 10, HSR_WORK_ARRAY(case_complete_p1_count, each_worker)
				Text x_pos+25, y_pos, 115, 10, HSR_WORK_ARRAY(worker_name_const, each_worker)
				If x_pos = 185 Then
					x_pos = 350
				Else
					x_pos = 185
					y_pos = y_pos + 10
				End If
			End If
		Next
		If x_pos = 350 Then y_pos = y_pos + 10
		y_pos = y_pos + 5
		If month_after_next_need_to_work <> 0 Then
			Text 180, y_pos, 130, 10, "Cases to Still Process in Phase 1: " & month_after_next_need_to_work
			Text 350, y_pos, 130, 10, "Percent: " & month_after_next_waiting_pcnt & " %"
		Else
			Text 180, y_pos, 260, 10, "All Ex Parte Evaluation cases for " & PHASE_ONE_MO & " have been completed."
		End If
		GroupBox 180, 75, 306, y_pos-75, "Count"
		Text 210, 75, 20, 10, "Name"
		Text 350, 75, 20, 10, "Count"
		Text 375, 75, 20, 10, "Name"

		y_pos = y_pos + 15
		' MsgBox "y_pos - " & y_pos
		If y_pos = 125 Then y_pos = 165
		GroupBox 5, 65, 485, y_pos-65, "PHASE ONE - " & PHASE_ONE_MO

		y_pos = y_pos + 25

		'PHASE 2
		set_y_pos = y_pos - 10
		Text 15, y_pos, 155, 10, "Total Cases with HC ER in " & PHASE_TWO_MO & ": " & next_month_er_count
		y_pos = y_pos + 15
		Text 15, y_pos, 175, 10, "Cases Ex Parte after Phase ONE: " & next_month_still_expt		'" - XX%"
		y_pos = y_pos + 10
		Text 95, y_pos, 175, 10, "Percent: " & next_month_initially_expt_pcnt & " %"
		y_pos = y_pos + 15
		Text 15, y_pos, 165, 10, "Cases with Phase 2 completed by HSR: " & next_month_hsr_phase2_complete_count
		y_pos = y_pos + 10
		Text 115, y_pos, 165, 10, "Percent: " & next_month_processed_pcnt & " %"
		y_pos = y_pos + 15
		If next_month_hsr_phase2_complete_count <> 0 Then
			Text 15, y_pos, 205, 10, "Cases Approved for " & PHASE_TWO_MO & ": " & next_month_app_count & "    ( " & next_month_app_pcnt & " % )"
			y_pos = y_pos + 10
			Text 15, y_pos, 205, 10, "Cases with ER Rescheduled : " & next_month_rescheduled_count & "    ( " & next_month_rescheduled_pcnt & " % )"
			y_pos = y_pos + 10
			Text 15, y_pos, 205, 10, "Cases closed/transferred : " & next_month_closed_xfer_count  & "    ( " & next_month_closed_xfer_pcnt & " % )"
			y_pos = y_pos + 10
			Text 15, y_pos, 205, 10, "Cases to on PROBLEM list: " & next_month_need_more_review & "    ( " & next_problem_pcnt & " % )"
		End If

		y_pos = set_y_pos + 10
		x_pos = 185
		For each_worker = 0 to UBound(HSR_WORK_ARRAY, 2)
			If HSR_WORK_ARRAY(case_complete_p2_count, each_worker) <> 0 Then
				If len(HSR_WORK_ARRAY(case_complete_p2_count, each_worker)) = 4 Then Text x_pos, y_pos, 15, 10, HSR_WORK_ARRAY(case_complete_p2_count, each_worker)
				If len(HSR_WORK_ARRAY(case_complete_p2_count, each_worker)) = 3 Then Text x_pos+5, y_pos, 15, 10, HSR_WORK_ARRAY(case_complete_p2_count, each_worker)
				If len(HSR_WORK_ARRAY(case_complete_p2_count, each_worker)) = 2 Then Text x_pos+10, y_pos, 15, 10, HSR_WORK_ARRAY(case_complete_p2_count, each_worker)
				If len(HSR_WORK_ARRAY(case_complete_p2_count, each_worker)) = 1 Then Text x_pos+15, y_pos, 15, 10, HSR_WORK_ARRAY(case_complete_p2_count, each_worker)
				Text x_pos+25, y_pos, 115, 10, HSR_WORK_ARRAY(worker_name_const, each_worker)
				If x_pos = 185 Then
					x_pos = 350
				Else
					x_pos = 185
					y_pos = y_pos + 10
				End If
			End If
		Next
		If x_pos = 350 Then y_pos = y_pos + 10
		y_pos = y_pos + 5
		If next_month_need_to_work <> 0 Then
			Text 180, y_pos, 130, 10, "Cases to Still Process in Phase 2: " & next_month_need_to_work
			Text 350, y_pos, 130, 10, "Percent: " & next_month_waiting_pcnt & " %"
		Else
			Text 180, y_pos, 260, 10, "All Ex Parte Approval cases for " & PHASE_TWO_MO & " have been completed."
		End If
		GroupBox 180, set_y_pos, 306, y_pos-set_y_pos, "Count"
		Text 210, set_y_pos, 20, 10, "Name"
		Text 350, set_y_pos, 20, 10, "Count"
		Text 375, set_y_pos, 20, 10, "Name"
		y_pos = y_pos + 15
		' MsgBox "y_pos - " & y_pos
		If y_pos < 310 then y_pos = 310
		If y_pos = 125 Then y_pos = 165
		GroupBox 5, set_y_pos-10, 485, y_pos-set_y_pos+10, "PHASE TWO - " & PHASE_TWO_MO

		ButtonGroup ButtonPressed
			OkButton 440, y_pos+5, 50, 15
	EndDialog

	Dialog Dialog1		'There is no looping and the dialog shows until the user presses OK or Cancel
	end_msg = ""
	Call script_end_procedure(end_msg)		'That's all in the ADMIN run
End If

bz_user = False
If user_ID_for_validation = "CALO001" Then bz_user = True
If user_ID_for_validation = "ILFE001" Then bz_user = True
If bz_user = False Then script_end_procedure("This script functionality can only be operated by the BlueZone Script Team. The script will now end.")

If ex_parte_function = "FIX LIST" Then
	Call script_end_procedure("There is no fix currently established.")
	fix_report_out = ""

	' 'THIS IS FOR CREATING A LIST OF CASES APPROVED FOR THE REVIEW MONTH THAT HAVE MSA and/or GRH
	' review_date = "8/1/2023"			'This sets a date as the review date to compare it to information in the data list and make sure it's a date
	' review_date = DateAdd("d", 0, review_date)

	' 'Opening a spreadsheet to capture the cases with a SMRT ending soon
	' Set ObjExcel = CreateObject("Excel.Application")
	' ObjExcel.Visible = True
	' Set objSMRTWorkbook = ObjExcel.Workbooks.Add()
	' ObjExcel.DisplayAlerts = True

	' 'Setting the first 4 col as worker, case number, name, and APPL date
	' ObjExcel.Cells(1, 1).Value = "CASE NUMBER"
	' ObjExcel.Cells(1, 2).Value = "APPROVAL WORKER"
	' ObjExcel.Cells(1, 3).Value = "MSA Status"
	' ObjExcel.Cells(1, 4).Value = "GRH Status"

	' FOR i = 1 to 8		'formatting the cells'
	' 	ObjExcel.Cells(1, i).Font.Bold = True		'bold font'
	' NEXT

	' excel_row = 2		'initializing the counter to move through the excel lines

	' objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"		'we only need to look at the cases for the specific review month

	' Set objConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
	' Set objRecordSet = CreateObject("ADODB.Recordset")

	' 'opening the connections and data table
	' objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	' objRecordSet.Open objSQL, objConnection

	' Do While NOT objRecordSet.Eof 					'Loop through each item on the CASE LIST Table
	' 	If objRecordSet("ExParteAfterPhase2") = "Approved as Ex Parte" Then
	' 		MAXIS_case_number = objRecordSet("CaseNumber") 		'SET THE MAXIS CASE NUMBER
	' 		Call back_to_SELF

	' 		Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)

	' 		If msa_case = True or grh_case = True Then
	' 			ObjExcel.Cells(excel_row, 1).Value = MAXIS_case_number
	' 			ObjExcel.Cells(excel_row, 2).Value = objRecordSet("Phase2HSR")
	' 			ObjExcel.Cells(excel_row, 3).Value = msa_status
	' 			ObjExcel.Cells(excel_row, 4).Value = grh_status
	' 			excel_row = excel_row + 1
	' 		End If
	' 	End If
	' 	objRecordSet.MoveNext			'now we go to the next case
	' Loop
	' objRecordSet.Close			'Closing all the data connections
	' objConnection.Close
	' Set objRecordSet=nothing
	' Set objConnection=nothing

	' For col_to_autofit = 1 to 4
	' 	ObjExcel.columns(col_to_autofit).AutoFit()
	' Next
	'----------------------------------------------------------------

	'This area is here in the event that we need to create an update process to the Ex Parte data list on a large number of cases.
	'This will need to be defined on a case-by-case scenario.

	Call script_end_procedure("Fix completed." & fix_report_out)

End If

'This is the first functionality to run after the data list is created. It will likely be run some time between the 6th and the 15th.
'This will evaluate the cases for Ex Parte, send the initial SVES QURY and create the other verification lists
If ex_parte_function = "Prep 1" Then
	review_date = ep_revw_mo & "/1/" & ep_revw_yr			'This sets a date as the review date to compare it to information in the data list and make sure it's a date
	review_date = DateAdd("d", 0, review_date)
	smrt_cut_off = DateAdd("m", 1, review_date)				'This is the cutoff date for SMRT ending to identify which ones we want to have evaluated

	'This functionality will remove any 'holds' with 'In Progress' marked. This is to make sure no cases get left behind if a script fails
	If reset_in_Progress = checked Then
		'This is opening the Ex Parte Case List data table so we can loop through it.
		objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"		'we only need to look at the cases for the specific review month

		Set objConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
		Set objRecordSet = CreateObject("ADODB.Recordset")

		'opening the connections and data table
		objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
		objRecordSet.Open objSQL, objConnection

		'Loop through each item on the CASE LIST Table
		Do While NOT objRecordSet.Eof
			If objRecordSet("PREP_Complete") = "In Progress" Then			'If the case is marked as 'In Progress' - we are going to remove it
				MAXIS_case_number = objRecordSet("CaseNumber") 				'SET THE MAXIS CASE NUMBER

				objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET PREP_Complete = '" & NULL & "'  WHERE CaseNumber = '" & MAXIS_case_number & "'"	'removing the 'In Progress' indicator and blanking it out

				Set objUpdateConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
				Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

				'opening the connections and data table
				objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
				objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection
			End If
			objRecordSet.MoveNext			'move to the next item in the table
		Loop
		objRecordSet.Close			'Closing all the data connections
		objConnection.Close
		Set objRecordSet=nothing
		Set objConnection=nothing
	End If

	va_count = 0		'initializing these counter variables at 0
	uc_count = 0
	rr_count = 0

	'Opening a spreadsheet to capture the cases with a SMRT ending soon
	Set ObjSMRTExcel = CreateObject("Excel.Application")
	ObjSMRTExcel.Visible = True
	Set objSMRTWorkbook = ObjSMRTExcel.Workbooks.Add()
	ObjSMRTExcel.DisplayAlerts = True

	'Setting the first 4 col as worker, case number, name, and APPL date
	ObjSMRTExcel.Cells(1, 1).Value = "CASE NUMBER"
	ObjSMRTExcel.Cells(1, 2).Value = "REF"
	ObjSMRTExcel.Cells(1, 3).Value = "NAME"
	ObjSMRTExcel.Cells(1, 4).Value = "PMI NUMBER"
	ObjSMRTExcel.Cells(1, 5).Value = "SSN"
	ObjSMRTExcel.Cells(1, 6).Value = "SMRT Cert End Date"
	ObjSMRTExcel.Cells(1, 7).Value = "SMRT Renewal Status"
	ObjSMRTExcel.Cells(1, 8).Value = "Ongoing SMRT"

	FOR i = 1 to 8		'formatting the cells'
		ObjSMRTExcel.Cells(1, i).Font.Bold = True		'bold font'
	NEXT

	smrt_excel_row = 2		'initializing the counter to move through the excel lines

	'This is opening the Ex Parte Case List data table so we can loop through it.
	objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"		'we only need to look at the cases for the specific review month

	Set objConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'opening the connections and data table
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	Do While NOT objRecordSet.Eof 					'Loop through each item on the CASE LIST Table
		'Pulling any case where the PREP_complete is null or blank
		If IsNull(objRecordSet("PREP_Complete")) = True or objRecordSet("PREP_Complete") = "" Then
			all_hc_is_ABD = ""				'resetting all these variables to blank at the beginning of each loop so information doesn't carry over from one case to another
			SSA_income_exists = ""
			JOBS_income_exists = ""
			VA_income_exists = ""
			BUSI_income_exists = ""
			case_has_no_income = ""
			case_has_EPD = ""

			appears_ex_parte = True			'we default this to true and find reasons that exclude the case from Ex Parte as we look at case data.
			all_hc_is_ABD = True
			case_has_EPD = False
			case_is_in_henn = False
			MAXIS_case_number = objRecordSet("CaseNumber") 		'SET THE MAXIS CASE NUMBER

			'Here we are setting the PREP_Complete to 'In Progress' to hold the case as being worked.
			'This portion of the script is required to be able to have more than one person operating the BULK run at the same time.
			objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET PREP_Complete = 'In Progress'  WHERE CaseNumber = '" & MAXIS_case_number & "'"

			Set objUpdateConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
			Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

			'opening the connections and data table
			objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
			objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection

			ReDim MEMBER_INFO_ARRAY(memb_last_const, 0)			'This is defined here without a preserve to blank it out at the beginning of every loop with a new case
			memb_count = 0										'resetting the counting variable to size the member array
			list_of_membs_on_hc = " "							'we need to keep a list members by pmi to know if a person is already accounted for as we find all the members and programs

			'We need to pull all of the instances from the ELIG table for the currently defined case number
			'This will list the HH member and eligibility program for HC. We will use this to start to determine if the case can be processed as Ex Parte
			objELIGSQL = "SELECT * FROM ES.ES_ExParte_EligList WHERE [CaseNumb] = '" & MAXIS_case_number & "'"

			Set objELIGConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
			Set objELIGRecordSet = CreateObject("ADODB.Recordset")

			'opening the connections and data table
			objELIGConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
			objELIGRecordSet.Open objELIGSQL, objELIGConnection

			person_found = False		'setting the default of if we have found a person in the list
			Do While NOT objELIGRecordSet.Eof
				list_of_membs_on_hc = list_of_membs_on_hc & objELIGRecordSet("PMINumber") & " "		'adding the PMI to the list of all PMIs known on the case
				person_found = True																	'indicating that there was a person in the list for this case
				memb_known = False																	'sets that we don't know if we have already looked at this person
				'now we loop through all of the people we have already found for this case - we only want 1 array instance per person.
				For known_membs = 0 to UBound(MEMBER_INFO_ARRAY, 2)
					If trim(objELIGRecordSet("PMINumber")) = MEMBER_INFO_ARRAY(memb_pmi_numb_const, known_membs) Then		'If the PMI matches one in the array, we are going to set the information to that array instance
						memb_known = True															'identifies that we know about this person and they are already in the array

						'figuring out which program type location the information should be saved in for this table data
						'each person on a case may have up to three different lines for different programs
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
						'This will read the ELIG type information and use it to identify if a case does NOT appear Ex Parte
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

						If appears_ex_parte = False AND objELIGRecordSet("EligType") <> "DP" Then all_hc_is_ABD = False		'identifying if the case has ABD basis or not
						If objELIGRecordSet("EligType") = "DP" Then case_has_EPD = True										'identifying if the case has MA-EPD
					End If
				Next

				'If this is an unknown member, and has not been added to the array already, we need to add it
				If memb_known = False Then
					ReDim Preserve MEMBER_INFO_ARRAY(memb_last_const, memb_count)								'resizing the array

					'setting personal information to the array
					MEMBER_INFO_ARRAY(memb_pmi_numb_const, memb_count) 	= trim(objELIGRecordSet("PMINumber"))
					MEMBER_INFO_ARRAY(memb_ssn_const, memb_count) 		= trim(objELIGRecordSet("SocialSecurityNbr"))
					name_var									 		= trim(objELIGRecordSet("Name"))		'we want to format the name corectly.
					name_array = split(name_var)
					MEMBER_INFO_ARRAY(memb_name_const, memb_count) = name_array(UBound(name_array))
					For name_item = 0 to UBound(name_array)-1
						MEMBER_INFO_ARRAY(memb_name_const, memb_count) = MEMBER_INFO_ARRAY(memb_name_const, memb_count) & " " & name_array(name_item)
					Next
					MEMBER_INFO_ARRAY(memb_active_hc_const, memb_count)	= True
					MEMBER_INFO_ARRAY(table_prog_1, memb_count) 		= trim(objELIGRecordSet("MajorProgram"))	'setting the program information
					MEMBER_INFO_ARRAY(table_type_1, memb_count) 		= trim(objELIGRecordSet("EligType"))

					'This will read the ELIG type information and use it to identify if a case does NOT appear Ex Parte
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

					If appears_ex_parte = False AND objELIGRecordSet("EligType") <> "DP" Then all_hc_is_ABD = False		'identifying if the case has ABD basis or not
					If objELIGRecordSet("EligType") = "DP" Then case_has_EPD = True										'identifying if the case has MA-EPD

					MEMBER_INFO_ARRAY(sql_rr_income_exists, memb_count) = False		'defaulting the income types for this case to false
					MEMBER_INFO_ARRAY(sql_va_income_exists, memb_count) = False
					MEMBER_INFO_ARRAY(sql_uc_income_exists, memb_count) = False

					memb_count = memb_count + 1		'incrementing the array counter up for the next loop
				End if
				objELIGRecordSet.MoveNext			'going to the next record
			Loop
			objELIGRecordSet.Close			'Closing all the data connections
			objELIGConnection.Close
			Set objELIGRecordSet=nothing
			Set objELIGConnection=nothing

			'If the ELIG types still indicate that the case is Ex Parte, we are going to check REVW to make sure the case meets renewal requirements
			If appears_ex_parte = True Then
				'check HC ER date in STAT/REVW
				Call navigate_to_MAXIS_screen_review_PRIV("STAT", "REVW", is_this_priv)
				If is_this_priv = True Then appears_ex_parte = False						'excluding cases that are privileged
				If is_this_priv = False Then
					Call write_value_and_transmit("X", 5, 71)
					EMReadScreen STAT_HC_ER_mo, 2, 8, 27
					EMReadScreen STAT_HC_ER_yr, 2, 8, 33
					If ep_revw_mo <> STAT_HC_ER_mo or ep_revw_yr <> STAT_HC_ER_yr Then  appears_ex_parte = False		'if this does not have the correct renewal month, we will exclude it from Ex Parte
				End If
			End If

			'If the case still appears Ex Parte, we are going to check if we are missing people, and check income for further determination of Ex Parte
			If appears_ex_parte = True Then
				'If we did not find people in the ELIG list, we are going to check ELIG/HC
				If person_found = False Then
					Call navigate_to_MAXIS_screen("STAT", "SUMM")		'Creating new ELIG results
					Call write_value_and_transmit("BGTX", 20, 71)
					Call MAXIS_background_check

					Call navigate_to_MAXIS_screen("ELIG", "HC  ")		'Navigate to ELIG/HC
					'Here we start at the top of ELIG/HC and read each row to find HC information
					hc_row = 8
					Do
						pers_type = ""		'blanking out variables so they don't carry over from loop to loop
						std = ""
						meth = ""
						waiv = ""

						'reading the main HC Elig information - member, program, status
						EMReadScreen read_ref_numb, 2, hc_row, 3
						EMReadScreen clt_hc_prog, 4, hc_row, 28
						EMReadScreen hc_prog_status, 6, hc_row, 50
						ref_row = hc_row
						Do while read_ref_numb = "  "				'this will read for the reference number if there are multiple programs for a single member
							ref_row = ref_row - 1
							EMReadScreen read_ref_numb, 2, ref_row, 3
						Loop

						If hc_prog_status = "ACTIVE" Then			'If HC is currently active, we need to read more details about the program/eligibility
							clt_hc_prog = trim(clt_hc_prog)			'formatting this to remove whitespace
							If clt_hc_prog <> "NO V" AND clt_hc_prog <> "NO R" and clt_hc_prog <> "" Then		'these are non-hc persons

								Call write_value_and_transmit("X", hc_row, 26)									'opening the ELIG detail spans
								If clt_hc_prog = "QMB" or clt_hc_prog = "SLMB" or clt_hc_prog = "QI1" Then		'If it is an MSP, we want to read the type only from a specific place
									elig_msp_prog = clt_hc_prog
									EMReadScreen pers_type, 2, 6, 56
								Else																			'These are MA type programs (not MSP)
									'Now we have to fund the current month in elig to get the current elig type
									col = 19
									Do
										EMReadScreen span_month, 2, 6, col										'reading the month in ELIG
										EMReadScreen span_year, 2, 6, col+3

										'if the span month matchest current month plus 1, we are going to grab elig from that month
										If span_month = MAXIS_footer_month and span_year = MAXIS_footer_year Then
											EMReadScreen pers_type, 2, 12, col - 2								'reading the ELIG TYPE
											EMReadScreen std, 1, 12, col + 3
											EMReadScreen meth, 1, 13, col + 2
											EMReadScreen waiv, 1, 17, col + 2
											Exit Do																'leaving once we've found the information for this elig
										End If
										col = col + 11			'this goes to the next column
									Loop until col = 85			'This is off the page - if we hit this, we did NOT find the elig type in this elig result

									'If we hit 85, we did not get the information. So we are going to read it from the last budget month (most current)
									If col = 85 Then
										EMReadScreen pers_type, 2, 12, 72										'reading the ELIG TYPE
										EMReadScreen std, 1, 12, 77
										EMReadScreen meth, 1, 13, 76
										EMReadScreen waiv, 1, 17, 76
									End If
								End If
								PF3			'leaving the elig detail information

								'now we need to add the information we just read to the member array
								memb_known = False										'default that the member know is false
								For known_membs = 0 to UBound(MEMBER_INFO_ARRAY, 2)								'Looking at all the members known in the array
									If MEMBER_INFO_ARRAY(memb_ref_numb_const, known_membs) = read_ref_numb Then	'if the member reference from ELIG matches the ARRAY, we are going to add more elig details
										memb_known = True														'look we found a person
										If MEMBER_INFO_ARRAY(table_prog_1, known_membs) = "" Then				'finding which area of the array is blank to save the elig information there
											MEMBER_INFO_ARRAY(table_prog_1, known_membs) 		= clt_hc_prog
											MEMBER_INFO_ARRAY(table_type_1, known_membs) 		= pers_type
										ElseIf MEMBER_INFO_ARRAY(table_prog_2, known_membs) = "" Then
											MEMBER_INFO_ARRAY(table_prog_2, known_membs) 		= clt_hc_prog
											MEMBER_INFO_ARRAY(table_type_2, known_membs) 		= pers_type
										ElseIf MEMBER_INFO_ARRAY(table_prog_3, known_membs) = "" Then
											MEMBER_INFO_ARRAY(table_prog_3, known_membs) 		= clt_hc_prog
											MEMBER_INFO_ARRAY(table_type_3, known_membs) 		= pers_type
										End If

										'This will read the ELIG type information and use it to identify if a case does NOT appear Ex Parte
										If clt_hc_prog = "EH" Then appears_ex_parte = False
										If pers_type = "AX" Then appears_ex_parte = False
										If pers_type = "AA" Then appears_ex_parte = False
										If pers_type = "DP" Then appears_ex_parte = False
										If pers_type = "CK" Then appears_ex_parte = False
										If pers_type = "CX" Then appears_ex_parte = False
										If pers_type = "CB" Then appears_ex_parte = False
										If pers_type = "CM" Then appears_ex_parte = False
										If pers_type = "13" Then appears_ex_parte = False 	'TYMA
										If pers_type = "14" Then appears_ex_parte = False 	'TYMA
										If pers_type = "09" Then appears_ex_parte = False 	'Adoption Assistance
										If pers_type = "11" Then appears_ex_parte = False 	'Auto Newborn
										If pers_type = "10" Then appears_ex_parte = False 	'Adoption Assistance
										If pers_type = "25" Then appears_ex_parte = False 	'Foster Care
										If pers_type = "PX" Then appears_ex_parte = False
										If pers_type = "PC" Then appears_ex_parte = False
										If pers_type = "BC" Then appears_ex_parte = False

										If appears_ex_parte = False AND pers_type <> "DP" Then all_hc_is_ABD = False		'identifying if the case has ABD basis or not
										If pers_type = "DP" Then case_has_EPD = True										'identifying if the case has MA-EPD
									End If
								Next

								'If this is an unknown member, and has not been added to the array already, we need to add it
								If memb_known = False Then
									ReDim Preserve MEMBER_INFO_ARRAY(memb_last_const, memb_count)								'resizing the array

									'setting personal information to the array
									MEMBER_INFO_ARRAY(memb_ref_numb_const, memb_count) = read_ref_numb
									MEMBER_INFO_ARRAY(memb_active_hc_const, memb_count)	= True
									MEMBER_INFO_ARRAY(table_prog_1, memb_count) 		= trim(clt_hc_prog)
									MEMBER_INFO_ARRAY(table_type_1, memb_count) 		= trim(pers_type)

									'This will read the ELIG type information and use it to identify if a case does NOT appear Ex Parte
									If clt_hc_prog = "EH" Then appears_ex_parte = False
									If pers_type = "AX" Then appears_ex_parte = False
									If pers_type = "AA" Then appears_ex_parte = False
									If pers_type = "DP" Then appears_ex_parte = False
									If pers_type = "CK" Then appears_ex_parte = False
									If pers_type = "CX" Then appears_ex_parte = False
									If pers_type = "CB" Then appears_ex_parte = False
									If pers_type = "CM" Then appears_ex_parte = False
									If pers_type = "13" Then appears_ex_parte = False 	'TYMA
									If pers_type = "14" Then appears_ex_parte = False 	'TYMA
									If pers_type = "09" Then appears_ex_parte = False 	'Adoption Assistance
									If pers_type = "11" Then appears_ex_parte = False 	'Auto Newborn
									If pers_type = "10" Then appears_ex_parte = False 	'Adoption Assistance
									If pers_type = "25" Then appears_ex_parte = False 	'Foster Care
									If pers_type = "PX" Then appears_ex_parte = False
									If pers_type = "PC" Then appears_ex_parte = False
									If pers_type = "BC" Then appears_ex_parte = False

									If appears_ex_parte = False AND pers_type <> "DP" Then all_hc_is_ABD = False		'identifying if the case has ABD basis or not
									If pers_type = "DP" Then case_has_EPD = True										'identifying if the case has MA-EPD

									MEMBER_INFO_ARRAY(sql_rr_income_exists, memb_count) = False		'defaulting the income types for this case to false
									MEMBER_INFO_ARRAY(sql_va_income_exists, memb_count) = False
									MEMBER_INFO_ARRAY(sql_uc_income_exists, memb_count) = False

									memb_count = memb_count + 1 	'incrementing the array counter up for the next loop
								End If

							End If
						End If
						hc_row = hc_row + 1												'now we go to the next row
						EMReadScreen next_ref_numb, 2, hc_row, 3						'read the next HC information to find when we've reeached the end of the list
						EMReadScreen next_maj_prog, 4, hc_row, 28
					Loop until next_ref_numb = "  " and next_maj_prog = "    "

					CALL back_to_SELF()													'going to STAT/MEMB - because there is misssing personal information for the members discovered in this way
					Do
						CALL navigate_to_MAXIS_screen("STAT", "MEMB")
						EMReadScreen memb_check, 4, 2, 48
					Loop until memb_check = "MEMB"

					at_least_one_hc_active = False										'this is a default to identify if HC is active on the case
					For known_membs = 0 to UBound(MEMBER_INFO_ARRAY, 2)					'loop through the member array
						Call write_value_and_transmit(MEMBER_INFO_ARRAY(memb_ref_numb_const, known_membs), 20, 76)		'navigate to the member for this instance of the array
						EMReadscreen last_name, 25, 6, 30								'read and cormat the name from MEMB
						EMReadscreen first_name, 12, 6, 63
						last_name = trim(replace(last_name, "_", "")) & " "
						first_name = trim(replace(first_name, "_", "")) & " "
						MEMBER_INFO_ARRAY(memb_name_const, known_membs) = first_name & " " & last_name
						EMReadScreen PMI_numb, 8, 4, 46									'capturing the PMI number
						PMI_numb = trim(PMI_numb)
						MEMBER_INFO_ARRAY(memb_pmi_numb_const, known_membs) = right("00000000" & PMI_numb, 8)			'we have to format the pmi to match the data list format (8 digits with leading 0s included)
						EMReadScreen MEMBER_INFO_ARRAY(memb_ssn_const, known_membs), 11, 7, 42							'catpturing the SSN
						MEMBER_INFO_ARRAY(memb_ssn_const, known_membs) = replace(MEMBER_INFO_ARRAY(memb_ssn_const, known_membs), " ", "")
						MEMBER_INFO_ARRAY(memb_ssn_const, known_membs) = replace(MEMBER_INFO_ARRAY(memb_ssn_const, known_membs), "_", "")
						If MEMBER_INFO_ARRAY(table_prog_1, known_membs) <> "" Then at_least_one_hc_active = True		'setting the variable that identifies there is HC active based on the ELIG read from HC/ELIG
						If MEMBER_INFO_ARRAY(table_prog_2, known_membs) <> "" Then at_least_one_hc_active = True
						If MEMBER_INFO_ARRAY(table_prog_3, known_membs) <> "" Then at_least_one_hc_active = True
						If MEMBER_INFO_ARRAY(table_prog_1, known_membs) <> "" or MEMBER_INFO_ARRAY(table_prog_2, known_membs) <> "" or MEMBER_INFO_ARRAY(table_prog_3, known_membs) <> "" Then
							list_of_membs_on_hc = list_of_membs_on_hc & MEMBER_INFO_ARRAY(memb_pmi_numb_const, known_membs) & " "		'adding individuals to our list of members on HC
						End If

					Next
					If at_least_one_hc_active = False Then appears_ex_parte = False			'if no one is on HC, this cannot be Ex Parte
				End If
			End If

			Call navigate_to_MAXIS_screen("STAT", "MEMB")		'now we go find all the HH members
			Call get_list_of_members

			'Now we are going to start looking at income information to remove any cases that have income thant disqualifies it from Ex parte
			SSA_income_exists = False				'setting these variables to false at the beginning of each loop through
			RR_income_exists = False
			VA_income_exists = False
			UC_income_exists = False
			PRISM_income_exists = False
			Other_UNEA_income_exists = False
			JOBS_income_exists = False
			BUSI_income_exists = False

			'Pulling all rows from the INCOME list for the case number we are currently processing
			objIncomeSQL = "SELECT * FROM ES.ES_ExParte_IncomeList WHERE [CaseNumber] = '" & MAXIS_case_number & "'"

			Set objIncomeConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
			Set objIncomeRecordSet = CreateObject("ADODB.Recordset")

			'opening the connections and data table
			objIncomeConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
			objIncomeRecordSet.Open objIncomeSQL, objIncomeConnection

			'looping through each row in this case
			Do While NOT objIncomeRecordSet.Eof
				income_for_person_is_on_HC = False			'default this variable to false, indicating if the income is for a person on HC
				If InStr(list_of_membs_on_hc, objIncomeRecordSet("PersonID")) <> 0 Then income_for_person_is_on_HC = True		'this compares the PMI for the income to the list of PMIS discovered in finding HC elig information

				'If this income is for someone on HC, we are going to assess the income detail to determine if the case should still be Ex Parte
				If income_for_person_is_on_HC = True Then
					If objIncomeRecordSet("IncExpTypeCode") = "UNEA" Then									'UNEA income exists each type code will set the boolean about that income typr for this case
						If objIncomeRecordSet("IncomeTypeCode") = "01" Then SSA_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "02" Then SSA_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "03" Then SSA_income_exists = True
						If objIncomeRecordSet("IncomeTypeCode") = "16" Then RR_income_exists = True
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
					If objIncomeRecordSet("IncExpTypeCode") = "JOBS" Then JOBS_income_exists = True					'we do not need to clarify further for JOBS or BUSI income, just indicate if these incomes exist.
					If objIncomeRecordSet("IncExpTypeCode") = "BUSI" Then BUSI_income_exists = True
				End If

				'Here we set if there is certain types of income on the case for any member. This will information the creation of verification lists
				For known_membs = 0 to UBound(MEMBER_INFO_ARRAY, 2)			'loop through all the members
					If trim(objIncomeRecordSet("PersonID")) = MEMBER_INFO_ARRAY(memb_pmi_numb_const, known_membs) Then						'if the PMI matches
						If objIncomeRecordSet("IncomeTypeCode") = "16" Then MEMBER_INFO_ARRAY(sql_rr_income_exists, known_membs) = True		'if the income type is any of the specified, identify that the income exists
						If objIncomeRecordSet("IncomeTypeCode") = "11" Then MEMBER_INFO_ARRAY(sql_va_income_exists, known_membs) = True
						If objIncomeRecordSet("IncomeTypeCode") = "12" Then MEMBER_INFO_ARRAY(sql_va_income_exists, known_membs) = True
						If objIncomeRecordSet("IncomeTypeCode") = "13" Then MEMBER_INFO_ARRAY(sql_va_income_exists, known_membs) = True
						If objIncomeRecordSet("IncomeTypeCode") = "38" Then MEMBER_INFO_ARRAY(sql_va_income_exists, known_membs) = True
						If objIncomeRecordSet("IncomeTypeCode") = "14" Then MEMBER_INFO_ARRAY(sql_uc_income_exists, known_membs) = True
					End If
				Next

				objIncomeRecordSet.MoveNext		'move to the next Income row
			Loop
			objIncomeRecordSet.Close			'Closing all the data connections
			objIncomeConnection.Close
			Set objIncomeRecordSet=nothing
			Set objIncomeConnection=nothing

			'This part is for logic to help us determine if the income impacts the Ex parte option
			case_has_no_income = False			'start at false for 'no income' basically false here means the case has income
			'If every income type is false, then the case has no income and the variable, 'case_has_no_income' is set to True, because it is true that there is no income.
			If SSA_income_exists = False and RR_income_exists = False and VA_income_exists = False and UC_income_exists = False and PRISM_income_exists = False and Other_UNEA_income_exists = False and JOBS_income_exists = False and BUSI_income_exists = False Then case_has_no_income = True

			'If the case apears Ex Parte at this point, we are going to do another assessment
			If appears_ex_parte = True Then
				'reading case program information and PW
				Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
				EMReadScreen case_pw, 7, 21, 14
				If left(case_pw, 4) = "X127" Then case_is_in_henn = True
				'This would exclude cases that are not in Hennepin or are Closed
				' If case_is_in_henn = False then  appears_ex_parte = False				'we are not going to exclude for inactive or out of county uuntil Phase 1 at this point
				' If case_active = False Then appears_ex_parte = False
				' If ma_status <> "ACTIVE" and msp_status <> "ACTIVE" Then appears_ex_parte = False

				'Any other UNEA, or JOBS/BUSI income requires the case be on SNAP or MFIP at this point
				If Other_UNEA_income_exists = True OR JOBS_income_exists = True OR BUSI_income_exists = True Then
					appears_ex_parte = False									'if there is JOBS/BUSI/Other UNEA - this cannot be ex parte
					If mfip_status = "ACTIVE" Then appears_ex_parte = True		'unless MFIP or SNAP is active
					If snap_status = "ACTIVE" Then appears_ex_parte = True
				End If
			End If

			'If the case still appears Ex Parte at this point, we need to start the verifications
			If appears_ex_parte = True Then
				'For each case that is indicated as potentially ExParte, we are going to take preperation actions
				last_va_count = va_count			'These are counting variables to set for each loop
				last_uc_count = uc_count
				last_rr_count = rr_count

				Call find_unea_information			'Now we are reading UNEA information for all the HH members

				Call back_to_SELF

				'Send a SVES/CURY for all persons on a case
				Call navigate_to_MAXIS_screen("INFC", "SVES")
				'checking for NON-DISCLOSURE AGREEMENT REQUIRED FOR ACCESS TO IEVS FUNCTIONS'
				EMReadScreen agreement_check, 9, 2, 24
				IF agreement_check = "Automated" THEN script_end_procedure("To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the screens and review the agreement.")

				'We need to loop through each HH Member on the case and send a QURY for every one.
				For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
					Call send_sves_qury("SSN", qury_finish)							'function to send a SVES/QURY
					MEMBER_INFO_ARRAY(sves_qury_sent, each_memb) = qury_finish		'set the output of the qury attempt to the member array

					'we are trying to find and update any rows in the INCOME list where the case number and pmi match exactly and the claim number is close to the SSN to se the QURY information
					objIncomeSQL = "UPDATE ES.ES_ExParte_IncomeList SET QURY_Sent = '" & qury_finish & "' WHERE [CaseNumber] = '" & MAXIS_case_number & "' and [PersonID] = '" & MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb) & "' and [ClaimNbr] like '" & MEMBER_INFO_ARRAY(memb_ssn_const, each_memb) & "%'"

					Set objIncomeConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
					Set objIncomeRecordSet = CreateObject("ADODB.Recordset")

					'opening the connections and data table
					objIncomeConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
					objIncomeRecordSet.Open objIncomeSQL, objIncomeConnection

					'If the member has HC eligibility, we need to check DISA to see if it is ending
					If MEMBER_INFO_ARRAY(memb_active_hc_const, each_memb)	= True Then
						call navigate_to_MAXIS_screen("STAT", "DISA")
						EMWriteScreen MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), 20, 76
						transmit

						EMReadScreen disa_vers, 1, 2, 78													'disa panel exists
						If disa_vers <> "0" Then
							EMReadScreen disa_hc_verif, 1, 13, 69											'disa is SMRT verified
							If disa_hc_verif = "2" Then
								EMReadScreen disa_cert_end_date, 10, 7, 69									'grabbing the cert period end date
								If disa_cert_end_date = "__ __ ____" Then disa_cert_end_date = ""
								disa_cert_end_date = replace(disa_cert_end_date, " ", "/")					'making this a date
								If disa_cert_end_date <> "" Then
									disa_cert_end_date  = DateAdd("d", 0, disa_cert_end_date)
									If DateDiff("d", disa_cert_end_date, smrt_cut_off) >=0 Then				'if the cert end date is before or equal to the smrt cut off set at the beginning of the run
										'set the SMRT information to the list of SMRT cases
										ObjSMRTExcel.Cells(smrt_excel_row, 1).Value = MAXIS_case_number
										ObjSMRTExcel.Cells(smrt_excel_row, 2).Value = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
										ObjSMRTExcel.Cells(smrt_excel_row, 3).Value = MEMBER_INFO_ARRAY(memb_name_const, each_memb)
										ObjSMRTExcel.Cells(smrt_excel_row, 4).Value = MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb)
										ObjSMRTExcel.Cells(smrt_excel_row, 5).Value = left(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 3) & "-" & mid(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4, 2) & "-" & right(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4)
										ObjSMRTExcel.Cells(smrt_excel_row, 6).Value = disa_cert_end_date
										smrt_excel_row = smrt_excel_row + 1									'counting to increment to the next excel row
									End If
								End If
							End If
						End If
					End If

					'If there is RR income listed from the SQL table and NOT from UNEA - it is going to save any member with RR income listed on SQL to the RR array for the verif list
					If MEMBER_INFO_ARRAY(sql_rr_income_exists, each_memb) = True and MEMBER_INFO_ARRAY(unea_RR_exists, each_memb) = False Then
						ReDim Preserve RR_INCOME_ARRAY(rr_last_const, rr_count)

						RR_INCOME_ARRAY(rr_case_numb_const, rr_count) = MAXIS_case_number
						RR_INCOME_ARRAY(rr_ref_numb_const, rr_count) = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
						RR_INCOME_ARRAY(rr_pers_name_const, rr_count) = MEMBER_INFO_ARRAY(memb_name_const, each_memb)
						RR_INCOME_ARRAY(rr_pers_ssn_const, rr_count) = left(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 3) & "-" & mid(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4, 2) & "-" & right(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4)
						RR_INCOME_ARRAY(rr_pers_pmi_const, rr_count) = MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb)
						RR_INCOME_ARRAY(rr_inc_type_code_const, rr_count) = income_type_code
						RR_INCOME_ARRAY(rr_inc_type_info_const, rr_count) = "Railroad Retirement"
						RR_INCOME_ARRAY(rr_claim_numb_const, rr_count) = ""
						RR_INCOME_ARRAY(rr_prosp_inc_const, rr_count) = "Unknown"

						rr_count = rr_count + 1
					End If

					'If there is VA income listed from the SQL table and NOT from UNEA - it is going to save any member with VA income listed on SQL to the VA array for the verif list
					If MEMBER_INFO_ARRAY(sql_va_income_exists, each_memb) = True and MEMBER_INFO_ARRAY(unea_VA_exists, each_memb) = False Then
						ReDim Preserve VA_INCOME_ARRAY(va_last_const, va_count)

						VA_INCOME_ARRAY(va_case_numb_const, va_count) = MAXIS_case_number
						VA_INCOME_ARRAY(va_ref_numb_const, va_count) = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
						VA_INCOME_ARRAY(va_pers_name_const, va_count) = MEMBER_INFO_ARRAY(memb_name_const, each_memb)
						VA_INCOME_ARRAY(va_pers_ssn_const, va_count) = left(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 3) & "-" & mid(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4, 2) & "-" & right(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4)
						VA_INCOME_ARRAY(va_pers_pmi_const, va_count) = MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb)
						VA_INCOME_ARRAY(va_inc_type_code_const, va_count) = ""
						VA_INCOME_ARRAY(va_inc_type_info_const, va_count) = "VA Income"
						VA_INCOME_ARRAY(va_claim_numb_const, va_count) = ""
						VA_INCOME_ARRAY(va_prosp_inc_const, va_count) = "Unknown"

						va_count = va_count + 1
					End If

					'If there is UC income listed from the SQL table and NOT from UNEA - it is going to save any member with UC income listed on SQL to the UC array for the verif list
					If MEMBER_INFO_ARRAY(sql_uc_income_exists, each_memb) = True and MEMBER_INFO_ARRAY(unea_UC_exists, each_memb) = False Then
						ReDim Preserve UC_INCOME_ARRAY(uc_last_const, uc_count)

						UC_INCOME_ARRAY(uc_case_numb_const, uc_count) = MAXIS_case_number
						UC_INCOME_ARRAY(uc_ref_numb_const, uc_count) = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
						UC_INCOME_ARRAY(uc_pers_name_const, uc_count) = MEMBER_INFO_ARRAY(memb_name_const, each_memb)
						UC_INCOME_ARRAY(uc_pers_ssn_const, uc_count) = left(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 3) & "-" & mid(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4, 2) & "-" & right(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4)
						UC_INCOME_ARRAY(uc_pers_pmi_const, uc_count) = MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb)
						UC_INCOME_ARRAY(uc_inc_type_code_const, uc_count) = ""
						UC_INCOME_ARRAY(uc_inc_type_info_const, uc_count) = "Unemployment"
						UC_INCOME_ARRAY(uc_claim_numb_const, uc_count) = ""
						UC_INCOME_ARRAY(uc_prosp_inc_const, uc_count) = "Unknown"

						uc_count = uc_count + 1
					End If
				Next

				'Now that we have an array saved, we are going to add it to the Excel sheet right away for UC, VA, or RR income.
				'We do it all at once because if we have a script error, this way we don't lose the information

				If va_count <> 0 Then								'if there is VA income found
					If last_va_count <> va_count Then				'and the va income found has incremented up since the last loop
						If last_va_count = 0 Then						'First, if this is the first VA income, we need to set the Excel sheet up
							va_excel_created = True						'Identifying that the VA excel list was created

							Set objVAExcel = CreateObject("Excel.Application")				'opening a new Excel sheet
							objVAExcel.Visible = True
							Set objVAWorkbook = objVAExcel.Workbooks.Add()
							objVAExcel.DisplayAlerts = True

							objVAExcel.Cells(1, 1).Value = "CASE NUMBER"					'Putting the headers in place for the Excel sheet
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

						'adding any va income from the array to the spreadsheet
						Do
							objVAExcel.Cells(va_excel_row, 1).value = VA_INCOME_ARRAY(va_case_numb_const, va_inc_count)
							objVAExcel.Cells(va_excel_row, 2).value = VA_INCOME_ARRAY(va_ref_numb_const, va_inc_count)
							objVAExcel.Cells(va_excel_row, 3).value = VA_INCOME_ARRAY(va_pers_name_const, va_inc_count)
							objVAExcel.Cells(va_excel_row, 4).value = VA_INCOME_ARRAY(va_pers_pmi_const, va_inc_count)
							objVAExcel.Cells(va_excel_row, 5).value = VA_INCOME_ARRAY(va_pers_ssn_const, va_inc_count)
							If VA_INCOME_ARRAY(va_inc_type_code_const, va_inc_count) <> "" Then objVAExcel.Cells(va_excel_row, 6).value = VA_INCOME_ARRAY(va_inc_type_code_const, va_inc_count) & " - " & VA_INCOME_ARRAY(va_inc_type_info_const, va_inc_count)
							If VA_INCOME_ARRAY(va_inc_type_code_const, va_inc_count) = "" Then objVAExcel.Cells(va_excel_row, 6).value = VA_INCOME_ARRAY(va_inc_type_info_const, va_inc_count)
							objVAExcel.Cells(va_excel_row, 7).value = VA_INCOME_ARRAY(va_claim_numb_const, va_inc_count)
							objVAExcel.Cells(va_excel_row, 8).value = VA_INCOME_ARRAY(va_prosp_inc_const, va_inc_count)

							va_inc_count = va_inc_count + 1			'going to the next array item
							va_excel_row = va_excel_row + 1			'going to the next row
						Loop until va_inc_count = va_count			'loop until the income count gets to the total of va counted
					End If
				End If

				If uc_count <> 0 Then								'If there is UC income found
					If last_uc_count <> uc_count Then				'and the UC income found has incremented up since the last loop
						If last_uc_count = 0 Then						'First, if this is the first UC income, we need to set the Excel sheet up
							uc_excel_created = True						'Identifying that the UC excel list was created

							Set objUCExcel = CreateObject("Excel.Application")				'opening a new Excel sheet
							objUCExcel.Visible = True
							Set objUCWorkbook = objUCExcel.Workbooks.Add()
							objUCExcel.DisplayAlerts = True

							objUCExcel.Cells(1, 1).Value = "CASE NUMBER"					'Putting the headers in place for the Excel sheet
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

						'adding any uc income from the array to the spreadsheet
						Do
							objUCExcel.Cells(uc_excel_row, 1).value = UC_INCOME_ARRAY(uc_case_numb_const, uc_inc_count)
							objUCExcel.Cells(uc_excel_row, 2).value = UC_INCOME_ARRAY(uc_ref_numb_const, uc_inc_count)
							objUCExcel.Cells(uc_excel_row, 3).value = UC_INCOME_ARRAY(uc_pers_name_const, uc_inc_count)
							objUCExcel.Cells(uc_excel_row, 4).value = UC_INCOME_ARRAY(uc_pers_pmi_const, uc_inc_count)
							objUCExcel.Cells(uc_excel_row, 5).value = UC_INCOME_ARRAY(uc_pers_ssn_const, uc_inc_count)
							If UC_INCOME_ARRAY(uc_inc_type_code_const, uc_inc_count) <> "" Then objUCExcel.Cells(uc_excel_row, 6).value = UC_INCOME_ARRAY(uc_inc_type_code_const, uc_inc_count) & " - " & UC_INCOME_ARRAY(uc_inc_type_info_const, uc_inc_count)
							If UC_INCOME_ARRAY(uc_inc_type_code_const, uc_inc_count) = "" Then objUCExcel.Cells(uc_excel_row, 6).value = UC_INCOME_ARRAY(uc_inc_type_info_const, uc_inc_count)
							objUCExcel.Cells(uc_excel_row, 7).value = UC_INCOME_ARRAY(uc_claim_numb_const, uc_inc_count)
							objUCExcel.Cells(uc_excel_row, 8).value = UC_INCOME_ARRAY(uc_prosp_inc_const, uc_inc_count)

							uc_inc_count = uc_inc_count + 1			'going to the next array item
							uc_excel_row = uc_excel_row + 1			'going to the next row
						Loop until uc_inc_count = uc_count			'loop until the income count gets to the total of uc counted
					End If
				End If


				If rr_count <> 0 Then								'If there is RR income found
					If last_rr_count <> rr_count Then				'and the RR income found has incremented up since the last loop
						If last_rr_count = 0 Then							'First, if this is the first RR income, we need to set the Excel sheet up
							rr_excel_created = True							'Identifying that the RR excel list was created

							Set objRRExcel = CreateObject("Excel.Application")				'opening a new Excel sheet
							objRRExcel.Visible = True
							Set objRRWorkbook = objRRExcel.Workbooks.Add()
							objRRExcel.DisplayAlerts = True

							objRRExcel.Cells(1, 1).Value = "CASE NUMBER"					'Putting the headers in place for the Excel sheet
							objRRExcel.Cells(1, 2).Value = "REF"
							objRRExcel.Cells(1, 3).Value = "NAME"
							objRRExcel.Cells(1, 4).Value = "PMI NUMBER"
							objRRExcel.Cells(1, 5).Value = "SSN"
							objRRExcel.Cells(1, 6).Value = "RR INC TYPE"
							objRRExcel.Cells(1, 7).Value = "RR CLAIM NUMB"
							objRRExcel.Cells(1, 8).Value = "CURR RR INCOME"
							objRRExcel.Cells(1, 9).Value = "Verified RR Income"
							objRRExcel.columns(2).NumberFormat = "@" 		'formatting as text

							FOR i = 1 to 9		'formatting the cells'
								objRRExcel.Cells(1, i).Font.Bold = True		'bold font'
							NEXT

							rr_excel_row = 2
							rr_inc_count = 0
						End If

						'adding any rr income from the array to the spreadsheet
						Do
							objRRExcel.Cells(rr_excel_row, 1).value = RR_INCOME_ARRAY(rr_case_numb_const, rr_inc_count)
							objRRExcel.Cells(rr_excel_row, 2).value = RR_INCOME_ARRAY(rr_ref_numb_const, rr_inc_count)
							objRRExcel.Cells(rr_excel_row, 3).value = RR_INCOME_ARRAY(rr_pers_name_const, rr_inc_count)
							objRRExcel.Cells(rr_excel_row, 4).value = RR_INCOME_ARRAY(rr_pers_pmi_const, rr_inc_count)
							objRRExcel.Cells(rr_excel_row, 5).value = RR_INCOME_ARRAY(rr_pers_ssn_const, rr_inc_count)
							If RR_INCOME_ARRAY(rr_inc_type_code_const, rr_inc_count) <> "" Then objRRExcel.Cells(rr_excel_row, 6).value = RR_INCOME_ARRAY(rr_inc_type_code_const, rr_inc_count) & " - " & RR_INCOME_ARRAY(rr_inc_type_info_const, rr_inc_count)
							If RR_INCOME_ARRAY(rr_inc_type_code_const, rr_inc_count) = "" Then objRRExcel.Cells(rr_excel_row, 6).value = RR_INCOME_ARRAY(rr_inc_type_info_const, rr_inc_count)
							objRRExcel.Cells(rr_excel_row, 7).value = RR_INCOME_ARRAY(rr_claim_numb_const, rr_inc_count)
							objRRExcel.Cells(rr_excel_row, 8).value = RR_INCOME_ARRAY(rr_prosp_inc_const, rr_inc_count)

							rr_inc_count = rr_inc_count + 1			'going to the next array item
							rr_excel_row = rr_excel_row + 1			'going to the next row
						Loop until rr_inc_count = rr_count			'loop until the income count gets to the total of rr counted
					End If
				End If
			End If

			Call back_to_SELF				'getting back to base

			'Now we are going to update the case list with the Ex parte evaluation done. This also removes the 'In Progress' marker
			prep_status = date														'prep status should be a date
			If appears_ex_parte = False Then prep_status = "Not Ex Parte"			'if this case is not ex parte, the prep status is reset

			'here is the update statement. setting the exparte eval and the income/case information for the case running
			objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET SelectExParte = '" & appears_ex_parte & "', PREP_Complete = '" & prep_status & "', AllHCisABD = '" & all_hc_is_ABD & "', SSAIncomExist = '" & SSA_income_exists & "', WagesExist = '" & JOBS_income_exists & "', VAIncomeExist = '" & VA_income_exists & "', SelfEmpExists = '" & BUSI_income_exists & "', NoIncome = '" & case_has_no_income & "', EPDonCase = '" & case_has_EPD & "' WHERE CaseNumber = '" & MAXIS_case_number & "'"

			Set objUpdateConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
			Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

			'opening the connections and data table
			objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
			objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection
		End If
		objRecordSet.MoveNext			'now we go to the next case
	Loop

	'now we format and save the verification lists
	For col_to_autofit = 1 to 9
		If va_excel_created = True Then objVAExcel.columns(col_to_autofit).AutoFit()
		If uc_excel_created = True Then objUCExcel.columns(col_to_autofit).AutoFit()
		If rr_excel_created = True Then objRRExcel.columns(col_to_autofit).AutoFit()

		ObjSMRTExcel.columns(col_to_autofit).AutoFit()
	Next

	If va_excel_created = True Then
		objVAExcel.ActiveSheet.ListObjects.Add(xlSrcRange, objVAExcel.Range("A1:I" & va_excel_row - 1), xlYes).Name = "Table1"
		objVAExcel.ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleMedium2"
		objVAExcel.ActiveWorkbook.SaveAs ex_parte_folder & "\VA Income Verifications\VA Income - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx"
	End If
	If uc_excel_created = True Then
		objUCExcel.ActiveSheet.ListObjects.Add(xlSrcRange, objUCExcel.Range("A1:I" & uc_excel_row - 1), xlYes).Name = "Table1"
		objUCExcel.ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleMedium2"
		objUCExcel.ActiveWorkbook.SaveAs ex_parte_folder & "\UC Income Verifications\UC Income - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx"
	End If
	If rr_excel_created = True Then
		objRRExcel.ActiveSheet.ListObjects.Add(xlSrcRange, objRRExcel.Range("A1:I" & rr_excel_row - 1), xlYes).Name = "Table1"
		objRRExcel.ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleMedium2"
		objRRExcel.ActiveWorkbook.SaveAs ex_parte_folder & "\RR Income Verifications\RR Income - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx"
	End If
	ObjSMRTExcel.ActiveSheet.ListObjects.Add(xlSrcRange, ObjSMRTExcel.Range("A1:H" & smrt_excel_row - 1), xlYes).Name = "Table1"
	ObjSMRTExcel.ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleMedium2"
	ObjSMRTExcel.ActiveWorkbook.SaveAs ex_parte_folder & "\SMRT Ending\SMRT Ending - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx"

    objRecordSet.Close			'Closing all the data connections
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing

	'We are going to set the display message for the end of the script run
	end_msg = "BULK Prep 1 Run has been completed."

	'declare the SQL statement that will query the database for all cases with the review month we are evaluating
	objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"

	Set objConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'opening the connections and data table
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	'Open The CASE LIST Table
	'Loop through each item on the CASE LIST Table
	case_count = 0
	ex_parte_count = 0
	Do While NOT objRecordSet.Eof
		case_count = case_count + 1															'counting all the cases
		If objRecordSet("SelectExParte") = True Then ex_parte_count = ex_parte_count + 1	'counting all the ex parte cases
		objRecordSet.MoveNext		'go to the next case
	Loop
    objRecordSet.Close			'Closing all the data connections
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing

	percent_ex_parte = ex_parte_count/case_count						'doing some calculations for see percentages
	percent_ex_parte = percent_ex_parte * 100
	percent_ex_parte = FormatNumber(percent_ex_parte, 2, -1, 0, -1)

	'Creating an end message to display the case list counts
	end_msg = end_msg & vbCr & "Cases appear to have a HC ER scheduled for " & ep_revw_mo & "/" & ep_revw_yr & ": " & case_count
	end_msg = end_msg & vbCr & "Cases that appear to meet Ex Parte Criteria: " & ex_parte_count
	end_msg = end_msg & vbCr & "This appears to be " & percent_ex_parte & "% of cases."

	'This is the end of the fucntionality and will just display the end message at the end of this script file.
End If

'This functionality will be run about 5 days after the first PREP run.
'This will read the SVES TPQY information that was received from the QURY in PREP 1 and update the STAT panels
If ex_parte_function = "Prep 2" Then
	'this is for testing - we want to know
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	If ObjFSO.FileExists(user_myDocs_folder & ep_revw_mo & "/" & ep_revw_yr & " - prep 2 sept second tpqy list.txt") Then
		Set objTextStream = ObjFSO.OpenTextFile(user_myDocs_folder & "ExParte Tracking Lists/" & ep_revw_mo & "-" & ep_revw_yr & " - prep 2 sept second tpqy list.txt", ForAppending, true)
	Else
		' MsgBox user_myDocs_folder & ep_revw_mo & "/" & ep_revw_yr & " - prep 2 sept second tpqy list.txt"
		Set objTextStream = ObjFSO.CreateTextFile(user_myDocs_folder & "ExParte Tracking Lists/" & ep_revw_mo & "-" & ep_revw_yr & " - prep 2 sept second tpqy list.txt", ForWriting, true)
	End If
	objTextStream.WriteLine "LIST START"

	review_date = ep_revw_mo & "/1/" & ep_revw_yr			'This sets a date as the review date to compare it to information in the data list and make sure it's a date
	review_date = DateAdd("d", 0, review_date)
	MsgBox "review_date - " & review_date

	'This functionality will remove any 'holds' with 'In Progress' marked. This is to make sure no cases get left behind if a script fails
	If reset_in_Progress = checked Then
		'This is opening the Ex Parte Case List data table so we can loop through it.
		objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"		'we only need to look at the cases for the specific review month

		Set objConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
		Set objRecordSet = CreateObject("ADODB.Recordset")

		'opening the connections and data table
		objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
		objRecordSet.Open objSQL, objConnection

		'Loop through each item on the CASE LIST Table
		Do While NOT objRecordSet.Eof
			If objRecordSet("PREP_Complete") = "In Progress" Then			'If the case is marked as 'In Progress' - we are going to remove it
				MAXIS_case_number = objRecordSet("CaseNumber") 				'SET THE MAXIS CASE NUMBER

				objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET PREP_Complete = '" & NULL & "'  WHERE CaseNumber = '" & MAXIS_case_number & "'"	'removing the 'In Progress' indicator and blanking it out

				Set objUpdateConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
				Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

				'opening the connections and data table
				objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
				objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection
			End If
			objRecordSet.MoveNext			'move to the next item in the table
		Loop
		objRecordSet.Close			'Closing all the data connections
		objConnection.Close
		Set objRecordSet=nothing
		Set objConnection=nothing
	End If

	'Opening a spreadsheet to capture the cases with a SMRT ending soon
	Set ObjExcel = CreateObject("Excel.Application")
	ObjExcel.Visible = True
	Set objWorkbook = ObjExcel.Workbooks.Add()
	ObjExcel.DisplayAlerts = True

	'Setting the first 4 col as worker, case number, name, and APPL date
	ObjExcel.Cells(1, 1).Value = "CASE NUMBER"
	ObjExcel.Cells(1, 2).Value = "REF"
	ObjExcel.Cells(1, 3).Value = "NAME"
	ObjExcel.Cells(1, 4).Value = "PMI NUMBER"
	ObjExcel.Cells(1, 5).Value = "SSN"
	ObjExcel.Cells(1, 6).Value = "Date of Death"

	FOR i = 1 to 6		'formatting the cells'
		ObjExcel.Cells(1, i).Font.Bold = True		'bold font'
	NEXT

	excel_row = 2		'initializing the counter to move through the excel lines

	yesterday = DateAdd("d", -1, date)

	'This is opening the Ex Parte Case List data table so we can loop through it.
	objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"

	Set objConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'opening the connections and data table
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	'Loop through each item on the CASE LIST Table
	Do While NOT objRecordSet.Eof
		If objRecordSet("SelectExParte") = True and objRecordSet("PREP_Complete") <> date and objRecordSet("PREP_Complete") <> yesterday Then
			'For each case that is indicated as Ex parte, we are going to update the case information
			MAXIS_case_number = objRecordSet("CaseNumber") 		'SET THE MAXIS CASE NUMBER

			'Here we are setting the PREP_Complete to 'In Progress' to hold the case as being worked.
			'This portion of the script is required to be able to have more than one person operating the BULK run at the same time.
			objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET PREP_Complete = 'In Progress'  WHERE CaseNumber = '" & MAXIS_case_number & "'"

			Set objUpdateConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
			Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

			'opening the connections and data table
			objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
			objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection

			'Here is functionality to be sure the case is able to be updated
			case_is_in_henn = False					'default this to false

			'reading case program information and PW
			Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
			EMReadScreen case_pw, 7, 21, 14									'reading the curent PW for the case
			If left(case_pw, 4) = "X127" Then case_is_in_henn = True		'identifying if the case is not in HENN
			kick_it_off_reason = ""											'create an explanation of why the case is being removed form the Ex Parte list
			If case_is_in_henn = False Then kick_it_off_reason = "Case not in 27"
			If case_active = False Then kick_it_off_reason = "Case not Active"
			If (case_active = False and case_pending = False and case_rein = False) or case_is_in_henn = False Then
				'WE ARE NOT going to update this here for now
				' select_ex_parte = False
				' objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET SelectExParte = '" & select_ex_parte & "', PREP_Complete = '" & kick_it_off_reason & "' WHERE CaseNumber = '" & MAXIS_case_number & "'"

				' Set objUpdateConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
				' Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

				' 'opening the connections and data table
				' objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
				' objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection
			Else
				ReDim MEMBER_INFO_ARRAY(memb_last_const, 0)							'Reset this array to blank at the beginning of each loop for each case.
				Do
					Call navigate_to_MAXIS_screen("STAT", "MEMB")					'making suyre we get to STAT MEMB
					EMReadScreen memb_check, 4, 2, 48
				Loop until memb_check = "MEMB"
				Call get_list_of_members											'get a list of all the HH memebers on the case

				'Read SVES/TPQY for all persons on a case
				For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)
					MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = False			'defaulting these to false
					MEMBER_INFO_ARRAY(tpqy_memb_has_rsdi, each_memb) = False
					MEMBER_INFO_ARRAY(tpqy_rsdi_has_disa, each_memb) = False

					Call navigate_to_MAXIS_screen("INFC", "SVES")					'navigate to SVES
					EMWriteScreen MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb), 5, 68				'Enter the PMI for the current member and open the TPQY
					Call write_value_and_transmit("TPQY", 20, 70)

					EMReadScreen check_TPQY_panel, 4, 2, 53 						'Reads for TPQY panel
					If check_TPQY_panel = "TPQY" Then
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_record, each_memb), 		1, 8, 39		'saving all tpqy information into the member array
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_record, each_memb), 		1, 8, 65
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb), 		10, 6, 61
						MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb) = trim(MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb))
						MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb), " ", "/")
						EMReadScreen sves_response, 8, 7, 22 		'Return Date
						sves_response = replace(sves_response," ", "/")
						' MsgBox "SSI record - " & MEMBER_INFO_ARRAY(tpqy_ssi_record, each_memb) & vbCr & "RSDI record - " & MEMBER_INFO_ARRAY(tpqy_rsdi_record, each_memb)
					End If
					transmit

					EMReadScreen check_BDXP_panel, 4, 2, 53 						'Reads fro BDXP panel
					If check_BDXP_panel = "BDXP" Then
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb), 	12, 5, 40		'saving all tpqy information into the member array
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb), 		12, 5, 69
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb), 	2, 6, 19
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb), 	8, 8, 16
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_net_amt, each_memb), 		8, 8, 32
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_railroad_ind, each_memb), 		1, 8, 69
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb), 	5, 11, 69
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb), 	5, 14, 69
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb), 	10, 15, 69
						MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb))
						MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb))
						MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb), " ", "")
						MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb))
						' MEMBER_INFO_ARRAY(tpqy_rsdi_staus_desc, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_staus_desc, each_memb))
						' MEMBER_INFO_ARRAY(tpqy_rsdi_paydate, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_paydate, each_memb))
						MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb))
						MEMBER_INFO_ARRAY(tpqy_rsdi_net_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_net_amt, each_memb))
						MEMBER_INFO_ARRAY(tpqy_railroad_ind, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_railroad_ind, each_memb))
						MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb))
						MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb))
						MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb) = Trim (MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb))
						MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb), " ", "/1/")
						MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb), " ", "/1/")
						MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb), " ", "/")
					End If
					transmit

					EMReadScreen check_BDXM_panel, 4, 2, 53 						'Reads for BDXM panel
					If check_BDXM_panel = "BDXM" Then
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_medi_claim_num, each_memb), 			13, 4, 29		'saving all tpqy information into the member array
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

					EMReadScreen check_SDXE_panel, 4, 2, 53 						'Reads for SDXE panel
					If check_SDXE_panel = "SDXE" Then
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_claim_numb, each_memb), 		12, 5, 36		'saving all tpqy information into the member array
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

					EMReadScreen check_SDXP_panel, 4, 2, 50 							'Reads for SDXP panel
					If check_SDXP_panel = "SDXP" Then
						EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_pay_date, each_memb), 			5, 4, 16		'saving all tpqy information into the member array
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

						If MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb) <> "C01" Then
							last_payment_date = ""
							sdx_row = 8
							Do
								EMReadScreen sdx_payment_type, 1, sdx_row, 25
								If sdx_payment_type <> "0" and sdx_payment_type <> " " Then
									EMReadScreen last_payment_date, 5, sdx_row, 3
									EMReadScreen last_payment_amt, 9, sdx_row, 13
									Exit Do
								End If
								sdx_row = sdx_row + 1
							Loop until sdx_payment_type = " "
							If last_payment_date <> "" Then
								MEMBER_INFO_ARRAY(tpqy_ssi_last_pay_date, each_memb) = replace(last_payment_date, " ", "/1/")
								MEMBER_INFO_ARRAY(tpqy_ssi_last_pay_date, each_memb) = DateAdd("d", 0, MEMBER_INFO_ARRAY(tpqy_ssi_last_pay_date, each_memb))
								MEMBER_INFO_ARRAY(tpqy_ssi_last_pay_amt, each_memb) = trim(last_payment_amt)
							End If
						End If

						MEMBER_INFO_ARRAY(tpqy_ssi_pay_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_pay_date, each_memb))
						MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb))
						If MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb) = "" Then MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb) = "0"
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
						MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = True
						MEMBER_INFO_ARRAY(tpqy_ssi_is_ongoing, each_memb)= False
						If MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb) = "C01" Then MEMBER_INFO_ARRAY(tpqy_ssi_is_ongoing, each_memb) = True

						' If MEMBER_INFO_ARRAY(tpqy_ssi_pay_code, each_memb) <> "C01" Then MsgBox "STOP"
					End If
					If MEMBER_INFO_ARRAY(tpqy_rsdi_record, each_memb) = "Y" Then
						If MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb) = "C" or MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb) = "E" Then
							MEMBER_INFO_ARRAY(tpqy_memb_has_rsdi, each_memb) = True
							If IsDate(MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb)) = True Then MEMBER_INFO_ARRAY(tpqy_rsdi_has_disa, each_memb) = True
						End If
					End If

					' MsgBox "MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb) - " & MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb) & vbCr & "MAXIS_case_number - " & MAXIS_case_number & vbCr & "sves_response - " & sves_response
					objIncomeSQL = "UPDATE ES.ES_ExParte_IncomeList SET TPQY_Response = '" & sves_response & "' WHERE [CaseNumber] = '" & MAXIS_case_number & "' and [PersonID] = '" & MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb) & "' and [QURY_Sent] != 'NULL'"

					Set objIncomeConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
					Set objIncomeRecordSet = CreateObject("ADODB.Recordset")

					'opening the connections and data table
					objIncomeConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
					objIncomeRecordSet.Open objIncomeSQL, objIncomeConnection

					Call back_to_SELF
				Next

				'navigating into STAT
				Do
					Call navigate_to_MAXIS_screen("STAT", "SUMM")
					EMReadScreen summ_check, 4, 2, 46
				Loop until summ_check = "SUMM"
				verif_types = ""						'blanking out the list of verifications for the CASE/NOTE

				'here we attempt to go update STAT with the information gathered from TPQY
				For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)									'looping thorugh each HH Member
					If IsDate(MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb)) = True Then 		'If there is a date of dealth listed, for now we are just going to add them to a list
						ObjExcel.Cells(excel_row, 1).Value = MAXIS_case_number
						ObjExcel.Cells(excel_row, 2).Value = MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
						ObjExcel.Cells(excel_row, 3).Value = MEMBER_INFO_ARRAY(memb_name_const, each_memb)
						ObjExcel.Cells(excel_row, 4).Value = MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb)
						ObjExcel.Cells(excel_row, 5).Value = left(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 3) & "-" & mid(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4, 2) & "-" & right(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4)
						ObjExcel.Cells(excel_row, 6).Value = MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb)
						excel_row = excel_row + 1													'counting to increment to the next excel row
					Else 	'If there is no date of death, we are going to try to update UNEA for SSI/RSDI
						'Update MAXIS UNEA panels with information from TPQY
						If MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = True Then				'Member with SSI
							If MEMBER_INFO_ARRAY(tpqy_ssi_is_ongoing, each_memb) = True Then		'If SSI appears to be ongoing (Current Pay)
								Call find_UNEA_panel(MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), "03", SSI_UNEA_instance, "", SSI_panel_found)

								Call update_unea_pane(SSI_panel_found, "03", MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb), MEMBER_INFO_ARRAY(tpqy_ssi_claim_numb, each_memb) & MEMBER_INFO_ARRAY(tpqy_ssi_recip_code, each_memb), MEMBER_INFO_ARRAY(tpqy_ssi_SSP_elig_date, each_memb), "", "")
								If InStr(verif_types, "SSI") = 0 Then verif_types = verif_types & "/SSI"
							ElseIf isDate(MEMBER_INFO_ARRAY(tpqy_ssi_last_pay_date, each_memb)) = True Then	'If SSI has an end date listed
								Call find_UNEA_panel(MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), "03", SSI_UNEA_instance, "", SSI_panel_found)

								Call update_unea_pane(SSI_panel_found, "03", MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb), MEMBER_INFO_ARRAY(tpqy_ssi_claim_numb, each_memb) & MEMBER_INFO_ARRAY(tpqy_ssi_recip_code, each_memb), MEMBER_INFO_ARRAY(tpqy_ssi_SSP_elig_date, each_memb), MEMBER_INFO_ARRAY(tpqy_ssi_last_pay_date, each_memb), MEMBER_INFO_ARRAY(tpqy_ssi_last_pay_amt, each_memb))
								If InStr(verif_types, "SSI End") = 0 Then verif_types = verif_types & "/SSI End"
							'There is no handling for if person appears to have SSI ended but we could not find an end date.
							End If
						End If

						If MEMBER_INFO_ARRAY(tpqy_memb_has_rsdi, each_memb) = True Then				'Member with RSDI
							'TODO - this functionality might need revision - not sure if the amount matching is the way to go
							If MEMBER_INFO_ARRAY(tpqy_rsdi_net_amt, each_memb) <> MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb) Then
								If MEMBER_INFO_ARRAY(tpqy_rsdi_has_disa, each_memb) = True Then
									rsdi_type = "01"
									Call find_UNEA_panel(MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), rsdi_type, RSDI_UNEA_instance, MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb), RSDI_panel_found)
								Else
									rsdi_type = "02"
									Call find_UNEA_panel(MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), rsdi_type, RSDI_UNEA_instance, MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb), RSDI_panel_found)
								End If
								Call update_unea_pane(RSDI_panel_found, rsdi_type, MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb), MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb), MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb), "", "")
								If InStr(verif_types, "RSDI") = 0 Then verif_types = verif_types & "/RSDI"
							End If
						End If

						'Update MAXIS MEDI panels with information from TPQY
						MEDI_panel_exists = False
						MEMBER_INFO_ARRAY(created_medi, each_memb) = False
						If MEMBER_INFO_ARRAY(tpqy_medi_claim_num, each_memb) <> "" Then

							EMWriteScreen "MEDI", 20, 71
							transmit
							EMReadScreen medi_check, 4, 2, 44
							Do while medi_check <> "MEDI"
								Call navigate_to_MAXIS_screen("STAT", "MEDI")
								EMReadScreen medi_check, 4, 2, 44
							Loop

							EMWriteScreen MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), 20, 76
							transmit

							EMReadScreen total_amt_of_panels, 1, 2, 78			'Checks to make sure there are JOBS panels for this member. If none exists, one will be created
							MEDI_panel_exists = True
							MEDI_active = False
							If total_amt_of_panels = "0" Then MEDI_panel_exists = False
							If (MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb) <> "" and MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb) = "") or (MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb) <> "" and MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb) = "") Then MEDI_active = True
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
								Loop Until begining_of_list = "FIRST PAGE"

								Do
									PF20
									EMReadScreen end_of_list, 9, 24, 14
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

									If row = 14 Then
										PF19
										EMReadScreen begining_of_list, 10, 24, 14

										If begining_of_list = "FIRST PAGE" Then
											Exit Do
										Else
											row = 17
										End If
									End If
								Loop

							End If
							If MEMBER_INFO_ARRAY(tpqy_part_a_start, each_memb) = "" and MEMBER_INFO_ARRAY(tpqy_part_a_stop, each_memb) = "" Then panel_part_a_accurate = True
							If MEMBER_INFO_ARRAY(tpqy_part_b_start, each_memb) = "" and MEMBER_INFO_ARRAY(tpqy_part_b_stop, each_memb) = "" Then panel_part_b_accurate = True

							If MEDI_panel_exists = True and (panel_part_a_accurate = False or panel_part_b_accurate = False) Then
								If InStr(verif_types, "Medicare") = 0 Then verif_types = verif_types & "/Medicare"
								PF9

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

										If row = 14 Then
											PF19
											EMReadScreen begining_of_list, 10, 24, 14

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

										If part_b_ended = True Then
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

										If row = 14 Then
											PF19
											EMReadScreen begining_of_list, 10, 24, 14
											If begining_of_list = "FIRST PAGE" Then
												Exit Do
											Else
												row = 17
											End If
										End If
									Loop
								End If
								transmit
							End If

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
								End If
							End If
						End If
					End If
				Next

				'Send the case through background
				Call write_value_and_transmit("BGTX", 20, 71)
				EMReadScreen wrap_check, 4, 2, 46
				If wrap_check = "WRAP" Then transmit
				Call back_to_SELF

				'here we are trying to update the INCOME List with the information found in TPQY
				For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)

					' If ssi_claim_numb <> "" or sves_rsdi_claim_numb <> "" Then
					If MEMBER_INFO_ARRAY(tpqy_ssi_claim_numb, each_memb) <> "" Then
						objIncUpdtSQL = "UPDATE ES.ES_ExParte_IncomeList SET TPQY_Response = '" & sves_response &_
										"', GrossAmt = '" & MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb) &_
										"', NetAmt = '" & "" &_
										"', EndDate = '" & NULL &_
										"' WHERE [PersonID] = '" & MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb) & "' and [ClaimNbr] = '" & MEMBER_INFO_ARRAY(tpqy_ssi_claim_numb, each_memb) & MEMBER_INFO_ARRAY(tpqy_ssi_recip_code, each_memb) & "'"

						Set objIncUpdtConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
						Set objIncUpdtRecordSet = CreateObject("ADODB.Recordset")

						'opening the connections and data table
						objIncUpdtConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
						objIncUpdtRecordSet.Open objIncUpdtSQL, objIncUpdtConnection

					End If
					If MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb) <> "" Then
						objIncUpdtSQL = "UPDATE ES.ES_ExParte_IncomeList SET TPQY_Response = '" & sves_response &_
										"', GrossAmt = '" & MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb) &_
										"', NetAmt = '" & MEMBER_INFO_ARRAY(tpqy_rsdi_net_amt, each_memb) &_
										"', EndDate = '" & MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb) &_
										"' WHERE [PersonID] = '" & MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb) & "' and [ClaimNbr] LIKE '" & left(MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb), 9) & "'"

						Set objIncUpdtConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
						Set objIncUpdtRecordSet = CreateObject("ADODB.Recordset")

						'opening the connections and data table
						objIncUpdtConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
						objIncUpdtRecordSet.Open objIncUpdtSQL, objIncUpdtConnection

					End If

					'If TPQY indicates that there may be a secondary clam number, we are going to send a QURY and save that to the INCOME list
					If MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb) <> "" Then
						Call send_sves_qury("CLAIM", qury_finish)
						MEMBER_INFO_ARRAY(second_qury_sent, each_memb) = qury_finish
						' MsgBox "qury_finish - " & qury_finish
						objIncUpdtSQL = "UPDATE ES.ES_ExParte_IncomeList SET QURY_Sent = '" & qury_finish & "' WHERE [CaseNumber] = '" & MAXIS_case_number & "' and [PersonID] = '" & MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb) & "' and [ClaimNbr] like '" & left(MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb), 9) & "%'"

						Set objIncUpdtConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
						Set objIncUpdtRecordSet = CreateObject("ADODB.Recordset")

						'opening the connections and data table
						objIncUpdtConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
						objIncUpdtRecordSet.Open objIncUpdtSQL, objIncUpdtConnection

						' MsgBox "check INC list - " & MAXIS_case_number
						objTextStream.WriteLine MAXIS_case_number & "| NAME: " & MEMBER_INFO_ARRAY(memb_name_const, each_memb) & "|" & "SSN: " &  MEMBER_INFO_ARRAY(memb_ssn_const, each_memb) & "| CLAIM NUMB: " & MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb) & "| QURY FINISH: " & qury_finish
					End If

				Next

				'CASE/NOTE details of the case information
				If left(verif_types, 1) = "/" Then verif_types = right(verif_types, len(verif_types)-1)
				note_title = "Verification of " & verif_types

				If verif_types <> "" Then
					Call navigate_to_MAXIS_screen("CASE", "NOTE")
					EMReadScreen last_note, 55, 5, 25
					EMReadScreen last_note_date, 8, 5, 6
					today_day = right("0"&DatePart("d", date), 2)
					today_mo = right("0"&DatePart("d", date), 2)
					today_yr = right(DatePart("d", date), 2)
					today_as_text = today_mo & "/" & today_day & "/" &today_yr

					last_note = trim(last_note)

					If last_note <> note_title or last_note_date <> today_as_text Then
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
									Call write_variable_in_CASE_NOTE(" * " & rsdi_inc & " Income of $ " & MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb) & " per month.")
								End If
								Call write_variable_in_CASE_NOTE("   - Verified through worker initiated data match.")
								Call write_variable_in_CASE_NOTE("   - UNEA panel updated eff " & MAXIS_footer_month & "/" & MAXIS_footer_year & ".")
							End If

							If IsDate(MEMBER_INFO_ARRAY(second_qury_sent, each_memb)) = True Then
								Call write_variable_in_CASE_NOTE("* Additional QURY sent for Claim numb: XXX-XX-" & right(MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb), len(MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb))-5))
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

			End If

			objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET PREP_Complete = '" & date & "' WHERE CaseNumber = '" & MAXIS_case_number & "'"

			Set objUpdateConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
			Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

			'opening the connections and data table
			objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
			objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection
		End If

		objRecordSet.MoveNext
	Loop
    objRecordSet.Close
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing

	ObjExcel.ActiveSheet.ListObjects.Add(xlSrcRange, ObjExcel.Range("A1:H" & excel_row - 1), xlYes).Name = "Table1"
	ObjExcel.ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleMedium2"
	ObjExcel.ActiveWorkbook.SaveAs ex_parte_folder & "\MEMBS with TPQY Date of Death - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx"


	end_msg = "BULK Prep 2 Run has been completed for " & review_date & "."


	'declare the SQL statement that will query the database
	objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"

	Set objConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'opening the connections and data table
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	'Open The CASE LIST Table
	'Loop through each item on the CASE LIST Table
	case_count = 0
	ex_parte_count = 0
	Do While NOT objRecordSet.Eof
		case_count = case_count + 1
		If objRecordSet("SelectExParte") = True Then ex_parte_count = ex_parte_count + 1
		If IsNull(objRecordSet("PREP_Complete")) = False Then prep_done_count = prep_done_count + 1
		If objRecordSet("PREP_Complete") = date Then prep_2_count = prep_2_count + 1
		If objRecordSet("PREP_Complete") = yesterday Then prep_2_count = prep_2_count + 1
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
	end_msg = end_msg & vbCr & vbCr & "Cases with PREP completed: " &  prep_done_count
	end_msg = end_msg & vbCr & "Cases where PREP 2 is completed: " & prep_2_count
End If

If ex_parte_function = "Phase 1" Then
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	tracking_doc_file = user_myDocs_folder & "ExParte Tracking Lists/Phase 1 " & ep_revw_mo & "/" & ep_revw_yr & " income update list.txt"
	If ObjFSO.FileExists(tracking_doc_file) Then
		Set objTextStream = ObjFSO.OpenTextFile(tracking_doc_file, ForAppending, true)
	Else
		Set objTextStream = ObjFSO.CreateTextFile(tracking_doc_file, ForWriting, true)
	End If
	objTextStream.WriteLine "LIST START"

	' prep_phase_2_run_date =
	va_excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\Ex Parte\VA Income Verifications\VA Income - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx"
	uc_excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\Ex Parte\UC Income Verifications\UC Income - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx"
	rr_excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\Ex Parte\RR Income Verifications\RR Income - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx"
	smrt_excel_file_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\Ex ParteSMRT Ending\SMRT Ending - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx"

	Do
		Do
			err_msg = ""
			Dialog1 = ""
			BeginDialog Dialog1, 0, 0, 400, 170, "Confirm Ex Parte process"
				Text 10, 15, 75, 10, "Date of Prep 2 Run:"
				EditBox 85, 10, 50, 15, prep_phase_2_run_date
				Text 10, 30, 75, 10, "Load VA List"
				EditBox 10, 40, 325, 15, va_excel_file_path
				Text 10, 60, 75, 10, "Load UC List"
				EditBox 10, 70, 325, 15, uc_excel_file_path
				Text 10, 90, 75, 10, "Load RR List"
				EditBox 10, 100, 325, 15, rr_excel_file_path
				Text 10, 120, 75, 10, "Load SMRT List"
				EditBox 10, 130, 325, 15, smrt_excel_file_path
				ButtonGroup ButtonPressed
					PushButton 185, 150, 210, 15, "Continue, all excel files are accurate.", continue_phase_1_btn
					' PushButton 230, 145, 100, 15, "Incorrect Process/Month", incorrect_process_btn
					PushButton 345, 40, 50, 15, "VA BROWSE", va_browse_btn
					PushButton 345, 70, 50, 15, "UC BROWSE", uc_browse_btn
					PushButton 345, 100, 50, 15, "RR BROWSE", rr_browse_btn
					PushButton 345, 130, 50, 15, "SMRT BROWSE", smrt_browse_btn
			EndDialog

			Dialog Dialog1
			cancel_without_confirmation

			' If IsDate(prep_phase_2_run_date) = False Then

		Loop until err_msg = ""
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	Loop until are_we_passworded_out = false					'loops until user passwords back in

	Call excel_open(uc_excel_file_path, True, True, ObjExcel, objWorkbook)

	uc_count = 0
	excel_row = 2

	Do

		ReDim Preserve UC_INCOME_ARRAY(uc_last_const, uc_count)

		UC_INCOME_ARRAY(uc_case_numb_const, uc_count) 	= ObjExcel.Cells(1, excel_row).Value
		UC_INCOME_ARRAY(uc_ref_numb_const, uc_count) 	= ObjExcel.Cells(2, excel_row).Value
		UC_INCOME_ARRAY(uc_pers_name_const, uc_count) 	= ObjExcel.Cells(3, excel_row).Value
		UC_INCOME_ARRAY(uc_pers_pmi_const, uc_count) 	= right("00000000" & timr(ObjExcel.Cells(4, excel_row).Value), 8)
		UC_INCOME_ARRAY(uc_pers_ssn_const, uc_count) 	= replace(trim(ObjExcel.Cells(5, excel_row).Value), "-", "")  'left(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 3) & "-" & mid(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4, 2) & "-" & right(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4)
		UC_INCOME_ARRAY(uc_inc_type_code_const, uc_count) = "14"
		UC_INCOME_ARRAY(uc_inc_type_info_const, uc_count) = "Unemployment"
		UC_INCOME_ARRAY(uc_claim_numb_const, uc_count) 	= ObjExcel.Cells(7, excel_row).Value
		UC_INCOME_ARRAY(uc_prosp_inc_const, uc_count) 	= ObjExcel.Cells(9, excel_row).Value
		UC_INCOME_ARRAY(uc_end_date_const, uc_count) 	= ObjExcel.Cells(10, excel_row).Value
		If IsNumeric(UC_INCOME_ARRAY(uc_prosp_inc_const, uc_count)) = True Then UC_INCOME_ARRAY(uc_prosp_inc_const, uc_count) = UC_INCOME_ARRAY(uc_prosp_inc_const, uc_count) * 1

		uc_count = uc_count + 1
		excel_row = excel_row + 1
		next_case_numb = ObjExcel.Cells(1, excel_row).Value
	Loop until next_case_numb = ""

	ObjExcel.ActiveWorkbook.Close

	ObjExcel.Application.Quit
	ObjExcel.Quit

	'Open VA verification spreadsheet and save to an array
	Call excel_open(uc_excel_file_path, True, True, ObjExcel, objWorkbook)

	va_count = 0
	excel_row = 2

	Do
		ReDim Preserve VA_INCOME_ARRAY(va_last_const, va_count)

		VA_INCOME_ARRAY(va_case_numb_const, va_count) 	= trim(ObjExcel.Cells(1, excel_row).Value)		'MAXIS_case_number
		VA_INCOME_ARRAY(va_ref_numb_const, va_count) 	= trim(ObjExcel.Cells(2, excel_row).Value)		'MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb)
		VA_INCOME_ARRAY(va_pers_name_const, va_count) 	= trim(ObjExcel.Cells(3, excel_row).Value)		'MEMBER_INFO_ARRAY(memb_name_const, each_memb)
		VA_INCOME_ARRAY(va_pers_pmi_const, va_count) 	= right("00000000" & trim(ObjExcel.Cells(4, excel_row).Value), 8)		'MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb)
		VA_INCOME_ARRAY(va_pers_ssn_const, va_count) 	= replace(trim(ObjExcel.Cells(5, excel_row).Value), "-", "")		'left(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 3) & "-" & mid(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4, 2) & "-" & right(MEMBER_INFO_ARRAY(memb_ssn_const, each_memb), 4)
		VA_INCOME_ARRAY(va_claim_numb_const, va_count) 	= trim(ObjExcel.Cells(7, excel_row).Value)
		VA_INCOME_ARRAY(va_prosp_inc_const, va_count) 	= trim(ObjExcel.Cells(9, excel_row).Value)
		If IsNumeric(VA_INCOME_ARRAY(va_prosp_inc_const, va_count)) = True Then VA_INCOME_ARRAY(va_prosp_inc_const, va_count) = VA_INCOME_ARRAY(va_prosp_inc_const, va_count) * 1

		va_type_from_excel = trim(ObjExcel.Cells(6, excel_row).Value)
		If InStr(va_type_from_excel, "-") = 0 Then
			VA_INCOME_ARRAY(va_inc_type_info_const, va_count) = va_type_from_excel
		Else
			temp_array = split(va_type_from_excel, "-")
			VA_INCOME_ARRAY(va_inc_type_code_const, va_count) = trim(temp_array(0))
			VA_INCOME_ARRAY(va_inc_type_info_const, va_count) = trim(temp_array(1))
		End If

		va_count = va_count + 1
		excel_row = excel_row + 1
		next_case_numb = ObjExcel.Cells(1, excel_row).Value
	Loop until next_case_numb = ""

	ObjExcel.ActiveWorkbook.Close

	ObjExcel.Application.Quit
	ObjExcel.Quit


	MAXIS_footer_month = CM_plus_1_mo
	MAXIS_footer_year = CM_plus_1_yr

	review_date = ep_revw_mo & "/1/" & ep_revw_yr
	review_date = DateAdd("d", 0, review_date)

	'This functionality will remove any 'holds' with 'In Progress' marked. This is to make sure no cases get left behind if a script fails
	If reset_in_Progress = checked Then
		'This is opening the Ex Parte Case List data table so we can loop through it.
		objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"		'we only need to look at the cases for the specific review month

		Set objConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
		Set objRecordSet = CreateObject("ADODB.Recordset")

		'opening the connections and data table
		objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
		objRecordSet.Open objSQL, objConnection

		'Loop through each item on the CASE LIST Table
		Do While NOT objRecordSet.Eof
		If objRecordSet("Phase1Complete") = "In Progress" Then			'If the case is marked as 'In Progress' - we are going to remove it
		MAXIS_case_number = objRecordSet("CaseNumber") 				'SET THE MAXIS CASE NUMBER

				objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET Phase1Complete = '" & NULL & "'  WHERE CaseNumber = '" & MAXIS_case_number & "'"	'removing the 'In Progress' indicator and blanking it out

				Set objUpdateConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
				Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

				'opening the connections and data table
				objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
				objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection
			End If
			objRecordSet.MoveNext			'move to the next item in the table
		Loop
		objRecordSet.Close			'Closing all the data connections
		objConnection.Close
		Set objRecordSet=nothing
		Set objConnection=nothing
	End If

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

	'Loop through each item on the CASE LIST Table
	Do While NOT objRecordSet.Eof
		If objRecordSet("SelectExParte") = True and (objRecordSet("Phase1Complete") = "" or IsNull(objRecordSet("Phase1Complete")) = True) Then
			kick_it_off_reason = ""
			case_active = ""
			case_is_in_henn = ""
			is_this_priv = ""

			'For each case that is indicated as Ex parte, we are going to update the case information
			MAXIS_case_number = objRecordSet("CaseNumber") 		'SET THE MAXIS CASE NUMBER

			'Here we are setting the Phase1Complete to 'In Progress' to hold the case as being worked.
			'This portion of the script is required to be able to have more than one person operating the BULK run at the same time.
			objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET Phase1Complete = 'In Progress'  WHERE CaseNumber = '" & MAXIS_case_number & "'"

			Set objUpdateConnection = CreateObject("ADODB.Connection")		'Creating objects for access to the SQL table
			Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

			'opening the connections and data table
			objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
			objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection

			case_is_in_henn = False
			Call navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", is_this_priv)
			If is_this_priv Then kick_it_off_reason = "PRIV case"
			If is_this_priv = False Then
				Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
				EMReadScreen case_pw, 7, 21, 14
				If left(case_pw, 4) = "X127" Then case_is_in_henn = True
				If case_is_in_henn = False Then kick_it_off_reason = "Case not in 27"
				If case_active = False Then kick_it_off_reason = "Case not Active"
			End If
			If kick_it_off_reason <> "" Then
				select_ex_parte = False
				objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET SelectExParte = '" & select_ex_parte & "', Phase1Complete = '" & kick_it_off_reason & "' WHERE CaseNumber = '" & MAXIS_case_number & "'"

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

					Call navigate_to_MAXIS_screen("INFC", "SVES")
					EMWriteScreen MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb), 5, 68
					EMWriteScreen "TPQY", 20, 70
					transmit

					EMReadScreen check_TPQY_panel, 4, 2, 53 		'Reads for TPQY panel
					If check_TPQY_panel = "TPQY" Then
						EMReadScreen tpqy_response_date, 8, 7, 22
						tpqy_response_date = trim(tpqy_response_date)
						If tpqy_response_date <> "" Then
							tpqy_response_date = replace(tpqy_response_date, " ", "/")
							tpqy_response_date = DateAdd("d", 0, tpqy_response_date)

							If DateDiff("d", prep_phase_2_run_date, tpqy_response_date) > 0 Then

								EMReadScreen tpqy_name_txt, 40, 4, 10
								EMReadScreen tpqy_ssn_txt, 11, 5, 9
								EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_record, each_memb), 		1, 8, 39
								EMReadScreen MEMBER_INFO_ARRAY(tpqy_ssi_record, each_memb), 		1, 8, 65
								EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb), 	12, 5, 35
								EMReadScreen MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb), 		10, 6, 61
								MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb) = trim(MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb))
								MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_date_of_death, each_memb), " ", "/")
								MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb))
								EMReadScreen sves_response, 8, 7, 22 		'Return Date
								sves_response = replace(sves_response," ", "/")
								objTextStream.WriteLine MAXIS_case_number & "| NAME: " & tpqy_name_txt & "|" & "SSN: " &  tpqy_ssn_txt & "|" & "SSI record - " & MEMBER_INFO_ARRAY(tpqy_ssi_record, each_memb) & "|" & "RSDI record - " & MEMBER_INFO_ARRAY(tpqy_rsdi_record, each_memb) & "| CLAIM NUMB: " & MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb)

								transmit

								EMReadScreen check_BDXP_panel, 4, 2, 53 		'Reads fro BDXP panel
								If check_BDXP_panel = "BDXP" Then
									EMReadScreen MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb), 		12, 5, 69
									EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb), 	2, 6, 19
									EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb), 	8, 8, 16
									EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_net_amt, each_memb), 		8, 8, 32
									EMReadScreen MEMBER_INFO_ARRAY(tpqy_railroad_ind, each_memb), 		1, 8, 69
									EMReadScreen MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb), 	5, 11, 69
									EMReadScreen MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb), 	5, 14, 69
									EMReadScreen MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb), 	10, 15, 69
									MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb))
									MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_dual_entl_nbr, each_memb), " ", "")
									MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_status_code, each_memb))
									' MEMBER_INFO_ARRAY(tpqy_rsdi_staus_desc, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_staus_desc, each_memb))
									' MEMBER_INFO_ARRAY(tpqy_rsdi_paydate, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_paydate, each_memb))
									MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb))
									MEMBER_INFO_ARRAY(tpqy_rsdi_net_amt, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_rsdi_net_amt, each_memb))
									MEMBER_INFO_ARRAY(tpqy_railroad_ind, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_railroad_ind, each_memb))
									MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb))
									MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb) = Trim(MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb))
									MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb) = Trim (MEMBER_INFO_ARRAY(tpqy_rsdi_disa_date, each_memb))
									MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb), " ", "/1/")
									MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb) = replace(MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb), " ", "/1/")
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
									If MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb) = "" Then MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb) = "0"
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

								' MsgBox "MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb) - " & MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb) & vbCr & "MAXIS_case_number - " & MAXIS_case_number & vbCr & "sves_response - " & sves_response
								objIncomeSQL = "UPDATE ES.ES_ExParte_IncomeList SET TPQY_Response = '" & sves_response & "' WHERE [CaseNumber] = '" & MAXIS_case_number & "' and [PersonID] = '" & MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb) & "' and [QURY_Sent] = '" & prep_phase_2_run_date & "'"

								'Creating objects for Access
								Set objIncomeConnection = CreateObject("ADODB.Connection")
								Set objIncomeRecordSet = CreateObject("ADODB.Recordset")

								'This is the file path for the statistics Access database.
								' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
								objIncomeConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
								objIncomeRecordSet.Open objIncomeSQL, objIncomeConnection


							End If
						End If
						Call back_to_SELF

					End If
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
					' If MEMBER_INFO_ARRAY(tpqy_memb_has_ssi, each_memb) = True Then
					' 	Call find_UNEA_panel(MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), "03", SSI_UNEA_instance, "", SSI_panel_found)

					' 	Call update_unea_pane(SSI_panel_found, "03", MEMBER_INFO_ARRAY(tpqy_ssi_gross_amt, each_memb), MEMBER_INFO_ARRAY(tpqy_ssi_claim_numb, each_memb) & MEMBER_INFO_ARRAY(tpqy_ssi_recip_code, each_memb), MEMBER_INFO_ARRAY(tpqy_ssi_SSP_elig_date, each_memb), "", "")
					' 	If InStr(verif_types, "SSI") = 0 Then verif_types = verif_types & "/SSI"
					' End If

					If MEMBER_INFO_ARRAY(tpqy_memb_has_rsdi, each_memb) = True Then
						If MEMBER_INFO_ARRAY(tpqy_rsdi_net_amt, each_memb) <> MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb) Then
							If MEMBER_INFO_ARRAY(tpqy_rsdi_has_disa, each_memb) = True Then
								rsdi_type = "01"
								Call find_UNEA_panel(MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), rsdi_type, RSDI_UNEA_instance, MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb), RSDI_panel_found)
							Else
								rsdi_type = "02"
								Call find_UNEA_panel(MEMBER_INFO_ARRAY(memb_ref_numb_const, each_memb), rsdi_type, RSDI_UNEA_instance, MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb), RSDI_panel_found)
							End If
							Call update_unea_pane(RSDI_panel_found, rsdi_type, MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb), MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb), MEMBER_INFO_ARRAY(tpqy_intl_entl_date, each_memb), "", "")
							If InStr(verif_types, "RSDI") = 0 Then verif_types = verif_types & "/RSDI"
						End If
					End If

				Next

				For each_uc = 0 to UBound(UC_INCOME_ARRAY)
					If UC_INCOME_ARRAY(uc_case_numb_const, each_uc) = MAXIS_case_number Then

						Call navigate_to_MAXIS_screen("STAT", "UNEA")
						EMWriteScreen UC_INCOME_ARRAY(uc_ref_numb_const, each_uc), 20, 76
						EMWriteScreen "01", 20, 79
						transmit

						Do
							EMReadScreen unea_inc_type, 2, 5, 37
							If unea_inc_type = "14" Then
								EMReadScreen unea_claim_number, 15, 6, 37
								unea_claim_number = replace(unea_claim_number, "_", "")

								If unea_claim_number = UC_INCOME_ARRAY(uc_claim_numb_const, each_uc) or UC_INCOME_ARRAY(uc_claim_numb_const, each_uc) = "" Then
									UC_INCOME_ARRAY(uc_panel_updated_const, each_uc) = "YES"
									PF9
									EMWriteScreen "6", 5, 65		'Write Other Verification Code "6"

									EMReadScreen first_pay_day, 8, 13, 54
									If first_pay_day <> "__ __ __" Then
										first_pay_day = replace(first_pay_day, " ", "/")
										first_pay_day = DateAdd("d", 0, first_pay_day)
										day_of_the_week = Weekday(first_pay_day)
									Else
										first_pay_day = ""
										day_of_the_week = 3
									End If

									first_of_cm_plus_1 = CM_plus_1_mo & "/1/" & CM_plus_1_yr
									first_of_cm_plus_1 = DateAdd("d", 0, first_of_cm_plus_1)
									first_of_cm_minus_1 = CM_minus_1_mo & "/1/" & CM_minus_1_yr
									first_of_cm_minus_1 = DateAdd("d", 0, first_of_cm_minus_1)
									Do While Weekday(first_of_cm_plus_1)<> day_of_the_week
										first_of_cm_plus_1 = DateAdd("d", 1, first_of_cm_plus_1)
									Loop
									Do While Weekday(first_of_cm_minus_1)<> day_of_the_week
										first_of_cm_minus_1 = DateAdd("d", 1, first_of_cm_minus_1)
									Loop

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

									If DateDiff("m", UC_INCOME_ARRAY(uc_end_date_const, each_uc), date) >=6 Then
										Call write_value_and_transmit("DEL", 20, 71)
									ElseIf UC_INCOME_ARRAY(uc_end_date_const, each_uc) <> "" Then
										Call create_mainframe_friendly_date(end_date, 7, 68, "YY")	'income end date (SSI: ssi_denial_date, RSDI: susp_term_date)
										Call write_value_and_transmit("X", 6, 56)
										Call clear_line_of_text(9, 65)
										Do
											transmit
											EMReadScreen HC_popup, 9, 7, 41
											' If HC_popup = "HC Income" then transmit
										Loop until HC_popup <> "HC Income"
									Else
										retro_date = first_of_cm_minus_1
										' MsgBox "retro_date - " & retro_date & vbCr & "start_date - " & start_date & vbCr & "DateDiff - " & DateDiff("d", retro_date, start_date)
										'TODO - this retro date thing failed
										EMReadScreen start_date, 8, 7, 37
										start_date = replace(start_date, " ", "/")
										start_date = DateAdd("d", 0, start_date)
										If DateDiff("d", retro_date, start_date) < 0 Then
											row = 13
											Do
												Call create_mainframe_friendly_date(retro_date, row, 25, "YY")
												EMWriteScreen UC_INCOME_ARRAY(uc_prosp_inc_const, each_uc), row, 39		'TODO: Testing values
												retro_date = DateAdd("w", 1, retro_date)
												row = row + 1
											Loop Until DateDiff("m", first_of_cm_minus_1, retro_date) = 1
										End If

										' MsgBox "STOP and look"
										prosp_date = first_of_cm_plus_1
										row = 13
										Do
											Call create_mainframe_friendly_date(prosp_date, row, 54, "YY")
											EMWriteScreen UC_INCOME_ARRAY(uc_prosp_inc_const, each_uc), row, 68		'TODO: Testing values
											prosp_date = DateAdd("w", 1, prosp_date)
											row = row + 1
										Loop Until DateDiff("m", first_of_cm_plus_1, prosp_date) = 1

										EMWriteScreen CM_plus_1_mo, 13, 54 'hardcoded dates
										EMWriteScreen "01", 13, 57
										EMWriteScreen CM_plus_1_yr, 13, 60 'hardcoded dates
										EMWriteScreen income_amount, 13, 68		'TODO: Testing values (income_amt which = rsdi_gross_amt or ssi_gross_amt )

										Call write_value_and_transmit("X", 6, 56)
										Call clear_line_of_text(9, 65)
										EMWriteScreen UC_INCOME_ARRAY(uc_prosp_inc_const, each_uc), 9, 65		'TODO: Testing values (rsdi_gross_amt or ssi_gross_amt )
										EMWriteScreen "4", 10, 63		'code for pay frequency
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

									If InStr(verif_types, "UC") = 0 Then verif_types = verif_types & "/UC"
									Exit Do
								End If
							End If
							transmit
							EMReadScreen last_unea, 7, 24, 2
						Loop until last_unea = "ENTER A"

					End If
				Next


				For each_va = 0 to UBound(VA_INCOME_ARRAY)
					If VA_INCOME_ARRAY(va_case_numb_const, eacheach_va_uc) = MAXIS_case_number Then

						Call navigate_to_MAXIS_screen("STAT", "UNEA")
						EMWriteScreen VA_INCOME_ARRAY(va_ref_numb_const, each_va), 20, 76
						transmit

						Do
							EMReadScreen unea_inc_type, 2, 5, 37
							If VA_INCOME_ARRAY(va_inc_type_code_const, va_count) = unea_inc_type or (VA_INCOME_ARRAY(va_inc_type_code_const, va_count) = "" and (unea_inc_type = "11" or unea_inc_type = "12" or unea_inc_type = "13" or unea_inc_type = "38")) Then
								If IsNumeric(VA_INCOME_ARRAY(va_prosp_inc_const, va_count)) = True Then
									VA_INCOME_ARRAY(va_panel_updated_const, va_count) = "YES"
									Call update_unea_pane(True, unea_inc_type, VA_INCOME_ARRAY(va_prosp_inc_const, va_count), VA_INCOME_ARRAY(va_claim_numb_const, va_count), "", "", "")
									If InStr(verif_types, "VA") = 0 Then verif_types = verif_types & "/VA"
									Exit Do
								Else
									PF9
									EMWriteScreen "N", 5, 65

									For unea_row = 13 to 17
										EMReadScreen pay_aount, 8, unea_row, 39
										If pay_aount <> "________" Then
											EMWriteScreen CM_minus_1_mo, unea_row, 25
											EMWriteScreen CM_minus_1_yr, unea_row, 31
										End If
									Next
									For unea_row = 13 to 17
										EMReadScreen pay_aount, 8, unea_row, 68
										If pay_aount <> "________" Then
											EMWriteScreen CM_minus_1_mo, unea_row, 54
											EMWriteScreen CM_minus_1_yr, unea_row, 60
										End If
									Next
								End If

							End If
							transmit
							EMReadScreen last_unea, 7, 24, 2
						Loop until last_unea = "ENTER A"
					End If
				Next





				'Send the case through background
				Call write_value_and_transmit("BGTX", 20, 71)
				EMReadScreen wrap_check, 4, 2, 46
				If wrap_check = "WRAP" Then transmit
				Call back_to_SELF

				For each_memb = 0 to UBound(MEMBER_INFO_ARRAY, 2)

					If MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb) <> "" Then
						objIncUpdtSQL = "UPDATE ES.ES_ExParte_IncomeList SET TPQY_Response = '" & sves_response &_
										"', GrossAmt = '" & MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb) &_
										"', NetAmt = '" & MEMBER_INFO_ARRAY(tpqy_rsdi_net_amt, each_memb) &_
										"', EndDate = '" & MEMBER_INFO_ARRAY(tpqy_susp_term_date, each_memb) &_
										"' WHERE [PersonID] = '" & MEMBER_INFO_ARRAY(memb_pmi_numb_const, each_memb) & "' and [ClaimNbr] LIKE '" & left(MEMBER_INFO_ARRAY(tpqy_rsdi_claim_numb, each_memb), 9) & "'"

						'Creating objects for Access
						Set objIncUpdtConnection = CreateObject("ADODB.Connection")
						Set objIncUpdtRecordSet = CreateObject("ADODB.Recordset")

						'This is the file path for the statistics Access database.
						' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
						objIncUpdtConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
						objIncUpdtRecordSet.Open objIncUpdtSQL, objIncUpdtConnection

					End If
				Next


				'CASE/NOTE details of the case information
				If left(verif_types, 1) = "/" Then verif_types = right(verif_types, len(verif_types)-1)
				note_title = "Verification of " & verif_types

				If verif_types <> "" Then
					Call navigate_to_MAXIS_screen("CASE", "NOTE")
					EMReadScreen last_note, 55, 5, 25
					EMReadScreen last_note_date, 8, 5, 6
					today_day = right("0"&DatePart("d", date), 2)
					today_mo = right("0"&DatePart("d", date), 2)
					today_yr = right(DatePart("d", date), 2)
					today_as_text = today_mo & "/" & today_day & "/" & today_yr

					last_note = trim(last_note)

					If last_note <> note_title or last_note_date <> today_as_text Then
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
									Call write_variable_in_CASE_NOTE(" * " & rsdi_inc & " Income of $ " & MEMBER_INFO_ARRAY(tpqy_rsdi_gross_amt, each_memb) & " per month.")
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
						For each_va = 0 to UBound(VA_INCOME_ARRAY)
							If VA_INCOME_ARRAY(uc_case_numb_const, each_va) = MAXIS_case_number Then
								If VA_INCOME_ARRAY(uc_panel_updated_const, each_va) = "YES"	Then
									Call write_variable_in_CASE_NOTE("Income from Unemployment for MEMB " & VA_INCOME_ARRAY(va_ref_numb_const, each_va) & " - " & VA_INCOME_ARRAY(uc_pers_name_const, each_va) & ".")
									Call write_variable_in_CASE_NOTE(" * Income of $ " & VA_INCOME_ARRAY(va_prosp_inc_const, each_va) & " per month.")
									objTextStream.WriteLine MAXIS_case_number & "| VA - MEMB: " & VA_INCOME_ARRAY(va_ref_numb_const, each_va)
								End If
							End If
						Next
						For each_uc = 0 to UBound(UC_INCOME_ARRAY)
							If UC_INCOME_ARRAY(uc_case_numb_const, each_uc) = MAXIS_case_number Then
								If UC_INCOME_ARRAY(uc_panel_updated_const, each_uc) = "YES"	Then
									Call write_variable_in_CASE_NOTE("Income from Unemployment for MEMB " & UC_INCOME_ARRAY(uc_ref_numb_const, each_uc) & " - " & UC_INCOME_ARRAY(uc_pers_name_const, each_uc) & ".")
									Call write_variable_in_CASE_NOTE(" * Income of $ " & UC_INCOME_ARRAY(uc_prosp_inc_const, each_uc) & " per week.")
									objTextStream.WriteLine MAXIS_case_number & "| UC - MEMB: " & UC_INCOME_ARRAY(uc_ref_numb_const, each_uc)
								End If
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

			End If
		End If
		objRecordSet.MoveNext
	Loop
	'Close the object
	objTextStream.Close

	objRecordSet.Close
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing

	end_msg = "BULK Phase 1 Run has been completed for " & review_date & "."



	'declare the SQL statement that will query the database
	objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"

	'Creating objects for Access
	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'This is the file path for the statistics Access database.
	' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	'Open The CASE LIST Table
	'Loop through each item on the CASE LIST Table
	case_count = 0
	ex_parte_count = 0
	phase_1_done_count = 0
	today_phase_1_count = 0
	cases_removed_from_ex_parte_in_phase_1 = 0
	Do While NOT objRecordSet.Eof
		case_count = case_count + 1
		If objRecordSet("SelectExParte") = True Then ex_parte_count = ex_parte_count + 1
		If IsNull(objRecordSet("Phase1Complete")) = False Then phase_1_done_count = phase_1_done_count + 1
		If objRecordSet("Phase1Complete") = date Then today_phase_1_count = today_phase_1_count + 1
		If objRecordSet("SelectExParte") = False and IsDate("PREP_Complete") = True Then cases_removed_from_ex_parte_in_phase_1 = cases_removed_from_ex_parte_in_phase_1 + 1
		objRecordSet.MoveNext
	Loop
    objRecordSet.Close
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing


	end_msg = end_msg & vbCr & "Cases appear to have a HC ER scheduled for " & ep_revw_mo & "/" & ep_revw_yr & ": " & case_count
	end_msg = end_msg & vbCr & "Cases that appear to meet Ex Parte Criteria: " & ex_parte_count & vbCr
	end_msg = end_msg & vbCr & "Cases that completed PREP but are NOT Ex Parte Now: " & cases_removed_from_ex_parte_in_phase_1 & vbCr

	end_msg = end_msg & vbCr & "Cases with Phase 1 Done: " & phase_1_done_count
	end_msg = end_msg & vbCr & "Cases with Phase 1 Done Today: " & today_phase_1_count
End If

If ex_parte_function = "Phase 2" Then
	' MsgBox "Phase 2 BULK Run Details to be added later. This functionality will prep cases for HSR Review at Phase 2, which will happen at the beginning of the Processing month (the month before the Review Month)."

	'TODO - add a check of the Phase 1 reason and information to update 'SelectExparte' if wrong
	Set ObjFSO = CreateObject("Scripting.FileSystemObject")
	tracking_doc_file = user_myDocs_folder & "ExParte Tracking Lists/Phase 2 " & ep_revw_mo & "/" & ep_revw_yr & " budg issues list.txt"
	If ObjFSO.FileExists(tracking_doc_file) Then
		Set objTextStream = ObjFSO.OpenTextFile(tracking_doc_file, ForAppending, true)
	Else
		Set objTextStream = ObjFSO.CreateTextFile(tracking_doc_file, ForWriting, true)
	End If
	objTextStream.WriteLine "LIST START"

	MAXIS_footer_month = CM_plus_1_mo
	MAXIS_footer_year = CM_plus_1_yr

	review_date = ep_revw_mo & "/1/" & ep_revw_yr
	review_date = DateAdd("d", 0, review_date)

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

	'Loop through each item on the CASE LIST Table
	Do While NOT objRecordSet.Eof
		' If objRecordSet("SelectExParte") = True and IsNull(objRecordSet("Phase1Complete")) = True Then
		process_this_one = False
		phase_2_complete = objRecordSet("Phase2Complete")
		If IsDate(phase_2_complete) = False then
			process_this_one = True
		Else
			If DateDiff("d", phase_2_complete, date) <> 0 and DateDiff("d", phase_2_complete, date) <> 1 Then process_this_one = True
		End If
		If objRecordSet("SelectExParte") = True and process_this_one = True Then
			MAXIS_case_number = objRecordSet("CaseNumber") 		'SET THE MAXIS CASE NUMBER

			Do
				Call navigate_to_MAXIS_screen("STAT", "SUMM")
				EMReadScreen summ_check, 4, 2, 46
			Loop until summ_check = "SUMM"
			verif_types = ""

			Call update_stat_budg

			'Send the case through background
			Call write_value_and_transmit("BGTX", 20, 71)
			EMReadScreen wrap_check, 4, 2, 46
			If wrap_check = "WRAP" Then transmit
			Call back_to_SELF

			objUpdateSQL = "UPDATE ES.ES_ExParte_CaseList SET Phase2Complete = '" & date & "' WHERE CaseNumber = '" & MAXIS_case_number & "'"
			'Creating objects for Access
			Set objUpdateConnection = CreateObject("ADODB.Connection")
			Set objUpdateRecordSet = CreateObject("ADODB.Recordset")

			'This is the file path for the statistics Access database.
			objUpdateConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
			objUpdateRecordSet.Open objUpdateSQL, objUpdateConnection

		End If
		objRecordSet.MoveNext
	Loop

	objRecordSet.Close
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing

	end_msg = "BULK Phase 2 Run has been completed for " & review_date & "."



	'declare the SQL statement that will query the database
	objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"

	'Creating objects for Access
	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'This is the file path for the statistics Access database.
	' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	'Open The CASE LIST Table
	'Loop through each item on the CASE LIST Table
	case_count = 0
	ex_parte_count = 0
	phase_2_done_count = 0
	today_phase_2_count = 0
	Do While NOT objRecordSet.Eof
		case_count = case_count + 1
		If objRecordSet("SelectExParte") = True Then ex_parte_count = ex_parte_count + 1
		If IsNull(objRecordSet("Phase2Complete")) = False Then phase_2_done_count = phase_2_done_count + 1
		If objRecordSet("Phase2Complete") = date Then today_phase_2_count = today_phase_2_count + 1
		objRecordSet.MoveNext
	Loop
    objRecordSet.Close
    objConnection.Close
    Set objRecordSet=nothing
    Set objConnection=nothing


	end_msg = end_msg & vbCr & "Cases appear to have a HC ER scheduled for " & ep_revw_mo & "/" & ep_revw_yr & ": " & case_count
	end_msg = end_msg & vbCr & "Cases that appear to meet Ex Parte Criteria: " & ex_parte_count & vbCr
	end_msg = end_msg & vbCr & "Cases with Phase 2 Done: " & phase_2_done_count
	end_msg = end_msg & vbCr & "Cases with Phase 2 Done Today: " & today_phase_2_count

End If

If ex_parte_function = "Check REVW information on Phase 1 Cases" Then
	'This should be run for cases at Phase 1 only after ER cutoff
	'Create a spreadsheet and pull all cases in the data table for CM + 2 review into the list
	'Opening a spreadsheet to capture the cases with a SMRT ending soon
	Set ObjExcel = CreateObject("Excel.Application")
	ObjExcel.Visible = True
	Set objSMRTWorkbook = ObjExcel.Workbooks.Add()
	ObjExcel.DisplayAlerts = True

	'Setting the first 4 col as worker, case number, name, and APPL date
	ObjExcel.Cells(1, 1).Value = "CASE NUMBER"
	ObjExcel.Cells(1, 2).Value = "Ex Parte WORKER"
	ObjExcel.Cells(1, 3).Value = "Select Ex Parte"
	ObjExcel.Cells(1, 4).Value = "Phase 1 Ex Parte Eval"
	ObjExcel.Cells(1, 5).Value = "Phase 1 Notes"

	ObjExcel.Cells(1, 6).Value = "REVW HC ER"
	ObjExcel.Cells(1, 7).Value = "REVW Status"
	ObjExcel.Cells(1, 8).Value = "Ex Parte Ind"
	ObjExcel.Cells(1, 9).Value = "Ex Parte REVW Mo"

	FOR i = 1 to 9		'formatting the cells'
		ObjExcel.Cells(1, i).Font.Bold = True		'bold font'
	NEXT


	'Capture the Ex Parte information from the table into the excel.
	excel_row = 2		'initializing the counter to move through the excel lines
	review_date = ep_revw_mo & "/1/" & ep_revw_yr
	review_date = DateAdd("d", 0, review_date)



	objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '" & review_date & "'"		'we only need to look at the cases for the specific review month

	Set objConnection = CreateObject("ADODB.Connection")	'Creating objects for access to the SQL table
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'opening the connections and data table
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	Do While NOT objRecordSet.Eof 					'Loop through each item on the CASE LIST Table
		MAXIS_case_number = objRecordSet("CaseNumber") 		'SET THE MAXIS CASE NUMBER
		Do
			If left(MAXIS_case_number, 1) = "0" Then MAXIS_case_number = right(MAXIS_case_number, len(MAXIS_case_number)-1)
		Loop until left(MAXIS_case_number, 1) <> "0"

		ObjExcel.Cells(excel_row, 1).Value = MAXIS_case_number
		ObjExcel.Cells(excel_row, 2).Value = objRecordSet("Phase1HSR")
		ObjExcel.Cells(excel_row, 3).Value = objRecordSet("SelectExParte")
		ObjExcel.Cells(excel_row, 4).Value = objRecordSet("ExParteAfterPhase1")
		ObjExcel.Cells(excel_row, 5).Value = objRecordSet("Phase1ExParteCancelReason")
		excel_row = excel_row + 1

		objRecordSet.MoveNext			'now we go to the next case
	Loop
	objRecordSet.Close			'Closing all the data connections
	objConnection.Close
	Set objRecordSet=nothing
	Set objConnection=nothing

	'Run to REPT/REVS for CM+2
	back_to_self    'We need to get back to SELF and manually update the footer month
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)

	Call navigate_to_MAXIS_screen("REPT", "REVS")
	EMWriteScreen ep_revw_mo, 20, 55
	EMWriteScreen ep_revw_yr, 20, 58
	transmit

	'Pull all REVS cases into an array
	const case_num_const 			= 0
	const hc_revw_status_const		= 1
	const hc_revw_er_month_const	= 2
	const hc_revw_ex_parte_yn_const	= 3
	const hc_revw_ex_parte_mo_const	= 4
	const hc_on_revs_const			= 5
	const case_found_on_sql			= 6
	const last_expt_const			= 7

	Dim EX_PARTE_REVW_INFO_ARRAY()
	ReDim EX_PARTE_REVW_INFO_ARRAY(last_expt_const, 0)

	case_count = 0

	'start of the FOR...next loop
	For each worker in worker_array
		worker = trim(worker)
		If worker = "" then exit for
		Call write_value_and_transmit(worker, 21, 6)   'writing in the worker number in the correct col

		'Grabbing case numbers from REVS for requested worker
		DO	'All of this loops until last_page_check = "THIS IS THE LAST PAGE"
			row = 7	'Setting or resetting this to look at the top of the list
			DO		'All of this loops until row = 19
				'Reading case information (case number, SNAP status, and cash status)
				EMReadScreen MAXIS_case_number, 8, row, 6
				MAXIS_case_number = trim(MAXIS_case_number)
				EmReadscreen HC_status, 1, row, 49

				'Navigates though until it runs out of case numbers to read
				IF MAXIS_case_number = "" then exit do

				'For some goofy reason the dash key shows up instead of the space key. No clue why. This will turn them into null variables.
				If HC_status = "-" 		then HC_status = ""

				If HC_status <> "" Then
					ReDim Preserve EX_PARTE_REVW_INFO_ARRAY(last_expt_const, case_count)

					EX_PARTE_REVW_INFO_ARRAY(case_num_const, case_count) = MAXIS_case_number
					EX_PARTE_REVW_INFO_ARRAY(hc_revw_status_const, case_count) = HC_status
					EX_PARTE_REVW_INFO_ARRAY(hc_on_revs_const, case_count) = True
					EX_PARTE_REVW_INFO_ARRAY(case_found_on_sql, case_count) = False

					case_count = case_count + 1
				End if

				row = row + 1    'On the next loop it must look to the next row
				MAXIS_case_number = "" 'Clearing variables before next loop
			Loop until row = 19		'Last row in REPT/REVS
			'Because we were on the last row, or exited the do...loop because the case number is blank, it PF8s, then reads for the "THIS IS THE LAST PAGE" message (if found, it exits the larger loop)
			PF8
			EMReadScreen last_page_check, 21, 24, 2	'checking to see if we're at the end
			'if max reviews are reached, the goes to next worker is applicable
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	next

	MAXIS_footer_month = CM_mo
	MAXIS_footer_year = CM_yr

	'navigate_to STAT to gather REVW information
	For revs_case = 0 to UBound(EX_PARTE_REVW_INFO_ARRAY, 2)
		MAXIS_case_number = EX_PARTE_REVW_INFO_ARRAY(case_num_const, revs_case)
		Call navigate_to_MAXIS_screen("STAT", "REVW")
		Call write_value_and_transmit("X", 5, 71)
		EMReadScreen HC_ER_Date, 8, 8, 27
		EMReadScreen ExPte_Ind, 1, 9, 27
		EMReadScreen ExPte_Mo, 7, 9, 71

		EX_PARTE_REVW_INFO_ARRAY(hc_revw_er_month_const, revs_case) = replace(HC_ER_Date, " ", "/")
		EX_PARTE_REVW_INFO_ARRAY(hc_revw_ex_parte_yn_const, revs_case) = ExPte_Ind
		EX_PARTE_REVW_INFO_ARRAY(hc_revw_ex_parte_mo_const, revs_case) = replace(ExPte_Mo, " ", "/")

		PF3
		Call back_to_SELF
	Next

	'Match the array cases to the ones on Excel and output the renewal information
	For xl_row = 2 to excel_row-1
		MAXIS_case_number = trim(ObjExcel.Cells(xl_row, 1).Value)
		case_found_on_revw = False
		For revs_case = 0 to UBound(EX_PARTE_REVW_INFO_ARRAY, 2)
			If MAXIS_case_number = EX_PARTE_REVW_INFO_ARRAY(case_num_const, revs_case) Then
				case_found_on_revw = True
				EX_PARTE_REVW_INFO_ARRAY(case_found_on_sql, revs_case) = True

				ObjExcel.Cells(xl_row, 6).Value = EX_PARTE_REVW_INFO_ARRAY(hc_revw_er_month_const, revs_case)
				ObjExcel.Cells(xl_row, 7).Value = EX_PARTE_REVW_INFO_ARRAY(hc_revw_status_const, revs_case)
				ObjExcel.Cells(xl_row, 8).Value = EX_PARTE_REVW_INFO_ARRAY(hc_revw_ex_parte_yn_const, revs_case)
				ObjExcel.Cells(xl_row, 9).Value = EX_PARTE_REVW_INFO_ARRAY(hc_revw_ex_parte_mo_const, revs_case)

			End If
		Next

		'navigate to STAT/REVW for any case that was not on the list.
		If case_found_on_revw = False Then
			Call navigate_to_MAXIS_screen("STAT", "REVW")
			EMReadScreen hc_revw_status, 1, 7, 73
			Call write_value_and_transmit("X", 5, 71)
			EMReadScreen HC_ER_Date, 8, 8, 27
			EMReadScreen ExPte_Ind, 1, 9, 27
			EMReadScreen ExPte_Mo, 7, 9, 71

			ObjExcel.Cells(xl_row, 6).Value = replace(HC_ER_Date, " ", "/")
			ObjExcel.Cells(xl_row, 7).Value = hc_revw_status
			ObjExcel.Cells(xl_row, 8).Value = ExPte_Ind
			ObjExcel.Cells(xl_row, 9).Value = replace(ExPte_Mo, " ", "/")

			PF3
			Call back_to_SELF
		End If
	Next
	'Add any cases that are in the REVS array to Excel if they were not already there
	For revs_case = 0 to UBound(EX_PARTE_REVW_INFO_ARRAY, 2)
		If EX_PARTE_REVW_INFO_ARRAY(case_found_on_sql, revs_case) = False Then
			ObjExcel.Cells(excel_row, 1).Value = EX_PARTE_REVW_INFO_ARRAY(case_num_const, revs_case)

			ObjExcel.Cells(excel_row, 6).Value = EX_PARTE_REVW_INFO_ARRAY(hc_revw_er_month_const, revs_case)
			ObjExcel.Cells(excel_row, 7).Value = EX_PARTE_REVW_INFO_ARRAY(hc_revw_status_const, revs_case)
			ObjExcel.Cells(excel_row, 8).Value = EX_PARTE_REVW_INFO_ARRAY(hc_revw_ex_parte_yn_const, revs_case)
			ObjExcel.Cells(excel_row, 9).Value = EX_PARTE_REVW_INFO_ARRAY(hc_revw_ex_parte_mo_const, revs_case)
			excel_row = excel_row + 1
		End If
	Next

	For col_to_autofit = 1 to 9
		ObjExcel.columns(col_to_autofit).AutoFit()
	Next

	objExcel.ActiveSheet.ListObjects.Add(xlSrcRange, objExcel.Range("A1:I" & excel_row - 1), xlYes).Name = "Table1"
	objExcel.ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleMedium2"
	objExcel.ActiveWorkbook.SaveAs ex_parte_folder & "\Phase 1 REVS Check - " & ep_revw_mo & "-" & ep_revw_yr & ".xlsx"

	Call script_end_procedure("We have a list of HC REVWs for " & ep_revw_mo & "/" & ep_revw_yr & ".")
End If

If ex_parte_function = "DHS Data Validation" Then
	data_sheet_file_path = "C:\Users\calo001\OneDrive - Hennepin County\Projects\Ex-Parte\Data Validation\DHS " & ep_revw_mo & ep_revw_yr & " List.xlsx"
	'ep_revw_mo
	'ep_revw_yr
	MAXIS_footer_month = CM_plus_1_mo
	MAXIS_footer_year = CM_plus_1_yr

	call excel_open(data_sheet_file_path, True, True, ObjExcel, objWorkbook)
	list_of_all_the_cases = " "


	PMI_01 = ""
	name_01 = ""
	person_01_ref_number = ""
	MAXIS_MA_prog_01 = ""
	MAXIS_MA_basis_01 = ""
	MAXIS_msp_prog_01 = ""
	name_02 = ""
	PMI_02 = ""
	person_02_ref_number = ""
	MAXIS_MA_prog_02 = ""
	MAXIS_MA_basis_02 = ""
	MAXIS_msp_prog_02 = ""

	excel_row = 3
	Do
		MAXIS_case_number = trim(ObjExcel.Cells(excel_row, 1).Value)
		MAXIS_case_number = right("00000000" & MAXIS_case_number, 8)
		list_of_all_the_cases = list_of_all_the_cases & MAXIS_case_number & " "

		Call read_boolean_from_excel(ObjExcel.Cells(excel_row, 15).Value, on_henn_list)
		Call read_boolean_from_excel(ObjExcel.Cells(excel_row, 16).Value, henn_appears_ex_parte)
		list_mx_maj_prog = trim(ObjExcel.Cells(excel_row, 7).Value)
		list_mx_msp = trim(ObjExcel.Cells(excel_row, 9).Value)

		If on_henn_list = True and henn_appears_ex_parte = True and list_mx_maj_prog = "" and list_mx_msp = "" Then
			found_on_sql = False
			sql_appears_ex_parte = False
			still_ex_parte = ""

			'declare the SQL statement that will query the database
			objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '2023-08-01'"
			' objSQL = "SELECT * FROM ES.ES_ExParte_CaseList"

			'Creating objects for Access
			Set objConnection = CreateObject("ADODB.Connection")
			Set objRecordSet = CreateObject("ADODB.Recordset")

			'This is the file path for the statistics Access database.
			' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
			objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
			objRecordSet.Open objSQL, objConnection

			Do While NOT objRecordSet.Eof
				sql_case_number = objRecordSet("CaseNumber")
				If MAXIS_case_number = sql_case_number Then
					found_on_sql = True
					sql_prep_complete = objRecordSet("PREP_Complete")
					still_ex_parte = objRecordSet("SelectExParte")
					If IsDate(sql_prep_complete) = True Then sql_appears_ex_parte = True

					objELIGSQL = "SELECT * FROM ES.ES_ExParte_EligList WHERE [CaseNumb] = '" & sql_case_number & "'"

					'Creating objects for Access
					Set objELIGConnection = CreateObject("ADODB.Connection")
					Set objELIGRecordSet = CreateObject("ADODB.Recordset")

					'This is the file path for the statistics Access database.
					' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
					objELIGConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
					objELIGRecordSet.Open objELIGSQL, objELIGConnection

					Do While NOT objELIGRecordSet.Eof

						If name_01 = "" Then
							name_01 = trim(objELIGRecordSet("Name"))
							PMI_01 = trim(objELIGRecordSet("PMINumber"))

							If objELIGRecordSet("MajorProgram") = "QM" or objELIGRecordSet("MajorProgram") = "SL"  Then
								MAXIS_msp_prog_01 = objELIGRecordSet("MajorProgram")
								MAXIS_msp_basis_01 = objELIGRecordSet("EligType")
							ElseIf objELIGRecordSet("MajorProgram") <> "MA" or objELIGRecordSet("MajorProgram") <> "IM" or objELIGRecordSet("MajorProgram") <> "EH" or objELIGRecordSet("MajorProgram") <> "NM"  Then
								MAXIS_MA_prog_01 = objELIGRecordSet("MajorProgram")
								MAXIS_MA_basis_01 = objELIGRecordSet("EligType")
							End If
						ElseIf PMI_01 = trim(objELIGRecordSet("PMINumber")) Then
							If objELIGRecordSet("MajorProgram") = "QM" or objELIGRecordSet("MajorProgram") = "SL"  Then
								MAXIS_msp_prog_01 = objELIGRecordSet("MajorProgram")
								MAXIS_msp_basis_01 = objELIGRecordSet("EligType")
							ElseIf objELIGRecordSet("MajorProgram") <> "MA" or objELIGRecordSet("MajorProgram") <> "IM" or objELIGRecordSet("MajorProgram") <> "EH" or objELIGRecordSet("MajorProgram") <> "NM"  Then
								MAXIS_MA_prog_01 = objELIGRecordSet("MajorProgram")
								MAXIS_MA_basis_01 = objELIGRecordSet("EligType")
							End If
						ElseIf name_02 = "" Then
							name_02 = trim(objELIGRecordSet("Name"))
							PMI_02 = trim(objELIGRecordSet("PMINumber"))

							If objELIGRecordSet("MajorProgram") = "QM" or objELIGRecordSet("MajorProgram") = "SL"  Then
								MAXIS_msp_prog_02 = objELIGRecordSet("MajorProgram")
								MAXIS_msp_basis_02 = objELIGRecordSet("EligType")
							ElseIf objELIGRecordSet("MajorProgram") <> "MA" or objELIGRecordSet("MajorProgram") <> "IM" or objELIGRecordSet("MajorProgram") <> "EH" or objELIGRecordSet("MajorProgram") <> "NM"  Then
								MAXIS_MA_prog_02 = objELIGRecordSet("MajorProgram")
								MAXIS_MA_basis_02 = objELIGRecordSet("EligType")
							End If
						ElseIf PMI_02 = trim(objELIGRecordSet("PMINumber")) Then
							If objELIGRecordSet("MajorProgram") = "QM" or objELIGRecordSet("MajorProgram") = "SL"  Then
								MAXIS_msp_prog_02 = objELIGRecordSet("MajorProgram")
								MAXIS_msp_basis_02 = objELIGRecordSet("EligType")
							ElseIf objELIGRecordSet("MajorProgram") <> "MA" or objELIGRecordSet("MajorProgram") <> "IM" or objELIGRecordSet("MajorProgram") <> "EH" or objELIGRecordSet("MajorProgram") <> "NM"  Then
								MAXIS_MA_prog_02 = objELIGRecordSet("MajorProgram")
								MAXIS_MA_basis_02 = objELIGRecordSet("EligType")
							End If
						End If
						objELIGRecordSet.MoveNext
					Loop

					ObjExcel.Cells(excel_row, 18).Value = PMI_01
					ObjExcel.Cells(excel_row, 19).Value = MAXIS_MA_prog_01
					ObjExcel.Cells(excel_row, 20).Value = MAXIS_MA_basis_01
					ObjExcel.Cells(excel_row, 21).Value = MAXIS_msp_prog_01
					ObjExcel.Cells(excel_row, 23).Value = PMI_02
					ObjExcel.Cells(excel_row, 24).Value = MAXIS_MA_prog_02
					ObjExcel.Cells(excel_row, 25).Value = MAXIS_MA_basis_02
					ObjExcel.Cells(excel_row, 26).Value = MAXIS_msp_prog_02

					objELIGRecordSet.Close
					objELIGConnection.Close
					Set objELIGRecordSet=nothing
					Set objELIGConnection=nothing



				End If
				objRecordSet.MoveNext
			Loop
			ObjExcel.Cells(excel_row, 15).Value = found_on_sql
			ObjExcel.Cells(excel_row, 16).Value = sql_appears_ex_parte
			ObjExcel.Cells(excel_row, 17).Value = still_ex_parte

			PMI_01 = ""
			name_01 = ""
			person_01_ref_number = ""
			MAXIS_MA_prog_01 = ""
			MAXIS_MA_basis_01 = ""
			MAXIS_msp_prog_01 = ""
			name_02 = ""
			PMI_02 = ""
			person_02_ref_number = ""
			MAXIS_MA_prog_02 = ""
			MAXIS_MA_basis_02 = ""
			MAXIS_msp_prog_02 = ""


			Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)

			ObjExcel.Cells(excel_row, 3).Value = ma_status
			ObjExcel.Cells(excel_row, 4).Value = msp_status

			Call navigate_to_MAXIS_screen("ELIG", "HC  ")
			hc_row = 8
			Do
				pers_type = ""
				std = ""
				meth = ""
				' elig_result = ""
				' results_created = ""
				waiv = ""
				EMReadScreen read_ref_numb, 2, hc_row, 3
				EMReadScreen clt_hc_prog, 4, hc_row, 28
				prev_row = hc_row
				Do while read_ref_numb = "  "
					prev_row = prev_row - 1
					EMReadScreen read_ref_numb, 2, prev_row, 3
				Loop


				clt_hc_prog = trim(clt_hc_prog)
				If clt_hc_prog <> "NO V" AND clt_hc_prog <> "NO R" and clt_hc_prog <> "" Then


					Call write_value_and_transmit("X", hc_row, 26)
					If clt_hc_prog = "QMB" or clt_hc_prog = "SLMB" or clt_hc_prog = "QI1" Then
						elig_msp_prog = clt_hc_prog
						EMReadScreen pers_type, 2, 6, 56
					Else
						col = 19
						Do									'Finding the current month in elig to get the current elig type
							EMReadScreen span_month, 2, 6, col
							EMReadScreen span_year, 2, 6, col+3

							If span_month = MAXIS_footer_month and span_year = MAXIS_footer_year Then		'reading the ELIG TYPE
								EMReadScreen pers_type, 2, 12, col - 2
								EMReadScreen std, 1, 12, col + 3
								EMReadScreen meth, 1, 13, col + 2
								EMReadScreen waiv, 1, 17, col + 2
								Exit Do
							End If
							col = col + 11
						Loop until col = 85
						If col = 85 Then
							Do
								col = col - 11
								EMReadScreen pers_type, 2, 12, col - 2
							Loop until pers_type <> "__" and pers_type <> "  "
						End If

					End If
					PF3

					If person_01_ref_number = "" Or person_01_ref_number = read_ref_numb Then
						person_01_ref_number = read_ref_numb
					ElseIf person_02_ref_number = "" Then
						person_02_ref_number = read_ref_numb
					End If


					If person_01_ref_number = read_ref_numb Then
						If clt_hc_prog = "QMB" or clt_hc_prog = "SLMB" or clt_hc_prog = "QI1" Then
							MAXIS_msp_prog_01 = clt_hc_prog
							MAXIS_msp_basis_01 = pers_type
						Else
							MAXIS_MA_prog_01 = clt_hc_prog
							MAXIS_MA_basis_01 = pers_type
						End If
					ElseIf person_02_ref_number = read_ref_numb Then
						If clt_hc_prog = "QMB" or clt_hc_prog = "SLMB" or clt_hc_prog = "QI1" Then
							MAXIS_msp_prog_02 = clt_hc_prog
							MAXIS_msp_basis_02 = pers_type
						Else
							MAXIS_MA_prog_02 = clt_hc_prog
							MAXIS_MA_basis_02 = pers_type
						End If
					End If
					' MsgBox "person_01_ref_number - " & person_01_ref_number & vbCr &_
					' 		"MAXIS_MA_prog_01 - " & MAXIS_MA_prog_01 & vbCr &_
					' 		"MAXIS_MA_basis_01 - " & MAXIS_MA_basis_01 & vbCr &_
					' 		"MAXIS_msp_prog_01 - " & MAXIS_msp_prog_01 & vbCr & vbCR &_
					' 		"person_02_ref_number - " & person_02_ref_number & vbCr &_
					' 		"MAXIS_MA_prog_02 - " & MAXIS_MA_prog_02 & vbCr &_
					' 		"MAXIS_MA_basis_02 - " & MAXIS_MA_basis_02 & vbCr &_
					' 		"MAXIS_msp_prog_02 - " & MAXIS_msp_prog_02
				End If
				hc_row = hc_row + 1
				EMReadScreen next_ref_numb, 2, hc_row, 3
				EMReadScreen next_maj_prog, 4, hc_row, 28
			Loop until next_ref_numb = "  " and next_maj_prog = "    "


			CALL back_to_SELF()
			If person_01_ref_number <> "" Then
				CALL navigate_to_MAXIS_screen("STAT", "MEMB")
				Do
					EMReadScreen read_ref_number, 2, 4, 33
					EMReadscreen last_name, 25, 6, 30
					EMReadscreen first_name, 12, 6, 63
					last_name = trim(replace(last_name, "_", "")) & " "
					first_name = trim(replace(first_name, "_", "")) & " "
					If read_ref_number = person_01_ref_number Then
						EMReadScreen PMI_01, 8, 4, 46
						PMI_01 = trim(PMI_01)
						PMI_01 = right("00000000" & PMI_01, 8)
						name_01 = first_name & " " & last_name
					End If
					If read_ref_number = person_02_ref_number Then
						EMReadScreen PMI_02, 8, 4, 46
						PMI_02 = trim(PMI_02)
						PMI_02 = right("00000000" & PMI_02, 8)
						name_02 = first_name & " " & last_name
					End If
					transmit
					EMReadScreen MEMB_end_check, 13, 24, 2
					' MsgBox "PMI_01 - " & PMI_01 & vbCr & "PMI_02 - " & PMI_02 & vbCr & "MEMB_end_check - " & MEMB_end_check
				LOOP Until PMI_01 <> "" AND (PMI_02 <> "" OR MEMB_end_check = "ENTER A VALID")
			End If

			ObjExcel.Cells(excel_row, 5).Value = person_01_ref_number
			ObjExcel.Cells(excel_row, 6).Value = PMI_01
			ObjExcel.Cells(excel_row, 7).Value = MAXIS_MA_prog_01
			ObjExcel.Cells(excel_row, 8).Value = MAXIS_MA_basis_01
			ObjExcel.Cells(excel_row, 9).Value = MAXIS_msp_prog_01
			ObjExcel.Cells(excel_row, 10).Value = person_02_ref_number
			ObjExcel.Cells(excel_row, 11).Value = PMI_02
			ObjExcel.Cells(excel_row, 12).Value = MAXIS_MA_prog_02
			ObjExcel.Cells(excel_row, 13).Value = MAXIS_MA_basis_02
			ObjExcel.Cells(excel_row, 14).Value = MAXIS_msp_prog_02

			PMI_01 = ""
			name_01 = ""
			person_01_ref_number = ""
			MAXIS_MA_prog_01 = ""
			MAXIS_MA_basis_01 = ""
			MAXIS_msp_prog_01 = ""
			name_02 = ""
			PMI_02 = ""
			person_02_ref_number = ""
			MAXIS_MA_prog_02 = ""
			MAXIS_MA_basis_02 = ""
			MAXIS_msp_prog_02 = ""
		End If

		excel_row = excel_row + 1
		next_case_number = trim(ObjExcel.Cells(excel_row, 1).Value)

	Loop Until next_case_number = ""

	'declare the SQL statement that will query the database
	objSQL = "SELECT * FROM ES.ES_ExParte_CaseList WHERE [HCEligReviewDate] = '2023-08-01'"
	' objSQL = "SELECT * FROM ES.ES_ExParte_CaseList"

	'Creating objects for Access
	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	'This is the file path for the statistics Access database.
	' stats_database_path = "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open objSQL, objConnection

	Do While NOT objRecordSet.Eof
		sql_case_number = objRecordSet("CaseNumber")
		If Instr(list_of_all_the_cases, sql_case_number) = 0 Then
		' If MAXIS_case_number = sql_case_number Then
			sql_appears_ex_parte = False
			' found_on_sql = True
			ObjExcel.Cells(excel_row, 1).Value = sql_case_number
			ObjExcel.Cells(excel_row, 2).Value = False
			ObjExcel.Cells(excel_row, 15).Value = True
			sql_prep_complete = objRecordSet("PREP_Complete")
			If IsDate(sql_prep_complete) = True Then sql_appears_ex_parte = True
			ObjExcel.Cells(excel_row, 16).Value = sql_appears_ex_parte
			excel_row = excel_row + 1

		End If
		objRecordSet.MoveNext
	Loop
	' ObjExcel.Cells(excel_row, 15).Value = found_on_sql

End If

'Loop through all the SQL Items and look for the right revew month and year and phase to determine if it's done.

Call script_end_procedure(end_msg)
