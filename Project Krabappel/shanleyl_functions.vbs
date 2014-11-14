Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\Users\Lucas\Documents\GitHub\DHS-MAXIS-Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'Excel Functions

Function excel_read(row,col)
	'Reads the value of the specified row and col in Excel.
	ObjExcel.Cells(row, col).Value
End Function

Function excel_write(row,col,cell_value)
	'Writes the cell_value variable in the specified row and col.	
	ObjExcel.Cells(row, col).Value = cell_value
End Function

Function excel_open_file(file_url)
	Set objWorkbook = objExcel.Workbooks.Open(file_url) 'Opens an excel file from a specific URL
	objExcel.DisplayAlerts = True
End Function

Function excel_open()
	Set objExcel = CreateObject("Excel.Application") 'Allows a user to perform functions within Microsoft Excel
	objExcel.Visible = True
End Function

'------------ MISC FUNCTIONS -----------------

Function add_months(D,E,F)
  'D = months to Add or Subtract 
  'E = Starting Date
  'F = Var to name the return variable
  calc_date = DateAdd("m", D, E)
  A = calc_date
  B = calc_date
  C = calc_date	
  A = Replace(Left(A, 2), "/", "")
  If len(A) = 1 then A = "0"&A
  If Mid(Mid(B, 3, 3),2,1) = "/" then
	B = "0"&Left(Mid(B, 3, 3),1)
  ElseIf Mid(Mid(B, 3, 3),2,1) <> "/" then
	B = Replace(Mid(B, 3, 3), "/", "")
	If len(B) = 1 then B = "0"&B
  End If
  C = Right(C,2)
  F = A & "/" & B & "/" & C
End Function

Function maxis_dater(A,B,C)
	'A = Input Date
	'B = Output Date Name
	'C = Specific name of date type
	
	error_message = "The date you used for your "& C &" is not a recognizable date format or was left blank."
	error_message_title = "Incorrect date format found."
	
	A = trim(A)
	A = replace(A, " ", "          ")
	A = replace(A, "  ", "          ")
	A = replace(A, "/", "          ")
	A = replace(A, "\", "          ")
	A = replace(A, "-", "          ")
	A = replace(A, ".", "          ")
	A = replace(A, ",", "          ")
	If InStr(A,"          ") = 0 then
		Do
			X = Len(A)
			If X < 4 or X > 8 then
				A = ""
				MsgBox error_message, error_message_title
				Exit Do				
			ElseIf X = 4 then
				If Left(A,2) = Left(year(now), 2) then
					A = ""
					MsgBox error_message, error_message_title
					Exit Do
				ElseIf Left(A,2) <> Left(year(now), 2) then
					valid_date = MsgBox("Did you mean: 0"&Left(A,1)&"/0"&Mid(A,2,1)&"/"&Right(A,2)&" for your "&C&"?",4,"Date Validation")
					If valid_date = 6 then 
						A = "0"&Left(A,1)&"/0"&Mid(A,2,1)&"/" & Right(A,2)
						Exit Do
					ElseIf valid_date = 7 then
						A = ""
						MsgBox error_message, error_message_title
						Exit Do
					End If
				End If
			ElseIf X = 5 then
				If Left(A,1) <> "0" then
					valid_date = MsgBox("Did you mean: 0"&Left(A,1)&"/"&Mid(A,2,2)&"/"&Right(A,2)&" for your "&C&"?",4,"Date Validation")
					If valid_date = 6 then 
						A = "0"&Left(A,1)&"/"&Mid(A,2,2)&"/"&Right(A,2)
						Exit Do
					End If
					If valid_date = 7 then valid_date = MsgBox("Did you mean: "&Left(A,2)&"/0"&Mid(A,3,1)&"/"&Right(A,2)&" for your "&C&"?",4,"Date Validation")
					If valid_date = 6 then 
						A = Left(A,2)&"/0"&Mid(A,3,1)&"/"&Right(A,2)
						Exit Do
					End If
					If valid_date = 7 then
						A = ""
						MsgBox error_message, error_message_title
						Exit Do
					End If
				ElseIf Left(A,1) = "0" then
					valid_date = MsgBox("Did you mean: "&Left(A,2)&"/0"&Mid(A,3,1)&"/"&Right(A,2)&" for your "&C&"?",4,"Date Validation")
					If valid_date = 6 then A = Left(A,2)&"/0"&Mid(A,3,1)&"/"&Right(A,2)
					If valid_date = 7 then
						A = ""
						MsgBox error_message, error_message_title
						Exit Do
					End If
				End If
			ElseIf X = 6 then
				If Left(A,1) <> "0" then
					valid_date = MsgBox("Did you mean: "&Left(A,2)&"/"&Mid(A,3,2)&"/"&Right(A,2)&" for your "&C&"?",4,"Date Validation")
					If valid_date = 6 then 
						A = Left(A,2)&"/"&Mid(A,3,2)&"/"&Right(A,2)
						Exit Do
					End If
					If valid_date = 7 then valid_date = MsgBox("Did you mean: 0"&Left(A,1)&"/0"&Mid(A,2,1)&"/"&Right(A,2)&" for your "&C&"?",4,"Date Validation")
					If valid_date = 6 then 
						A = "0"&Left(A,1)&"/0"&Mid(A,2,1)&"/"&Right(A,2)
						Exit Do
					End If
					If valid_date = 7 then
						A = ""
						MsgBox error_message, error_message_title
						Exit Do
					End If
				ElseIf Left(A,1) = "0" then
					valid_date = MsgBox("Did you mean: "&Left(A,2)&"/"&Mid(A,3,2)&"/"&Right(A,2)&" for your "&C&"?",4,"Date Validation")
					If valid_date = 6 then A = Left(A,2)&"/"&Mid(A,3,2)&"/"&Right(A,2)
					If valid_date = 7 then
						A = ""
						MsgBox error_message, error_message_title
						Exit Do
					End If
				End If
			ElseIf X = 7 then
				If Left(A,1) <> "0" then
					valid_date = MsgBox("Did you mean: 0"&Left(A,1)&"/"&Mid(A,2,2)&"/"&Right(A,2)&" for your "&C&"?",4,"Date Validation")
					If valid_date = 6 then 
						A = "0"&Left(A,1)&"/"&Mid(A,2,2)&"/"&Right(A,2)
						Exit Do
					End If
					If valid_date = 7 then valid_date = MsgBox("Did you mean: "&Left(A,2)&"/0"&Mid(A,3,1)&"/"&Right(A,2)&" for your "&C&"?",4,"Date Validation")
					If valid_date = 6 then 
						A = Left(A,2)&"/0"&Mid(A,3,1)&"/"&Right(A,2)
						Exit Do
					End If
					If valid_date = 7 then
						A = ""
						MsgBox error_message, error_message_title
						Exit Do
					End If
				ElseIf Left(A,1) = "0" then
					valid_date = MsgBox("Did you mean: "&Left(A,2)&"/0"&Mid(A,3,1)&"/"&Right(A,2)&" for your "&C&"?",4,"Date Validation")
					If valid_date = 6 then A = Left(A,2)&"/0"&Mid(A,3,1)&"/"&Right(A,2)
					If valid_date = 7 then
						A = ""
						MsgBox error_message, error_message_title
						Exit Do
					End If
				End If
			ElseIf X = 8 then
				valid_date = MsgBox("Did you mean: "&Left(A,2)&"/"&Mid(A,3,2)&"/"&Right(A,2)&" for your "&C&"?",4,"Date Validation")
				If valid_date = 6 then 
					A = Left(A,2)&"/"&Mid(A,3,2)&"/"&Right(A,2)
					Exit Do
				End If
				If valid_date = 7 then
					A = ""
					MsgBox error_message,error_message_title
					Exit Do
				End If
			End If
		Loop until valid_date = 6
	ElseIf InStr(A,"          ") <> 0 then 
		X = trim(Left(A, 5))
		If len(X) = 1 then X = "0" & X
		Y = trim(Mid(A, 5, 10))
		If len(Y) = 1 then Y = "0" & Y
		Z = trim(Right(A, 5))	
		If len(Z) = 4 then Z = Right(Z, 2)
		B = X & "/" & Y & "/" & Z
	End If	
End Function

Function debugger()
	EMReadScreen panel_error_found, 80, 24, 1
	panel_error_found = trim(panel_error_found)
	If panel_error_found <> "" then
		msgbox("Error Found: " & panel_error_found)
	End If
End Function

'--------------- PANEL FUNCTIONS ---------------------

Function write_panel_to_maxis_MSUR(msur_begin_date)
	call navigate_to_screen("STAT","MSUR")
	call create_panel_if_nonexistent
	call maxis_dater(msur_begin_date,return_date,"MNSure Begin Date")
	'msur_begin_date This is the date MSUR began for this client  
	col = 36
	row = 7
	'Places the Begin Date on the next available line	
	'This Can be uncommented once End Date logic is in place------------
	
	'Do
	'	msgbox row
	'	EMReadScreen no_date_check, 2, row, col	
	'	If no_date_check <> "__" then
	'		row = row + 2
	'	End If
	'Loop until no_date_check = "__"
	
	'-------------------------------------------------------------------
	EMWriteScreen left(return_date, 2)			, row, col		'Enters Month
	EMWriteScreen mid(return_date, 4, 2)		, row, col + 3	'Enters Day
	EMWriteScreen "20" & right(return_date, 2)	, row, col + 6	'Enters Year
	transmit
End Function

Function write_panel_to_maxis_PDED(pded_wid_deduction,pded_adult_child_disregard,pded_wid_disregard,pded_unea_income_deduction_reason,pded_unea_income_deduction_value,pded_earned_income_deduction_reason,pded_earned_income_deduction_value,pded_ma_epd_inc_asset_limit,pded_guard_fee,pded_rep_payee_fee,pded_other_expense,pded_shel_spcl_needs,pded_excess_need,pded_restaurant_meals)
	call navigate_to_screen("STAT","PDED")
	call create_panel_if_nonexistent
	
	'====================== FIX LATER ============================
	
	'Pickle Disregard
		'ADD ME LATER
		
	'=============================================================
		
	'Disa Widow/ers Deductionpded_shel_spcl_needs
	If pded_wid_deduction <> "" then
		pded_wid_deduction = ucase(pded_wid_deduction)
		pded_wid_deduction = left(pded_wid_deduction,1)
		EMWriteScreen pded_wid_deduction, 7, 60
	End If
	
	'Disa Adult Child Disregard
	If pded_adult_child_disregard <> "" then
		pded_adult_child_disregard = ucase(pded_adult_child_disregard)
		pded_adult_child_disregard = left(pded_adult_child_disregard,1)
		EMWriteScreen pded_adult_child_disregard, 8, 60
	End If
	
	'Widow/ers Disregard
	If pded_wid_disregard <> "" then
		pded_wid_disregard = ucase(pded_wid_disregard)
		pded_wid_disregard = left(pded_wid_disregard,1)
		EMWriteScreen pded_wid_disregard, 9, 60
	End If

	'Other Unearned Income Deduction
	If pded_unea_income_deduction_reason <> "" and pded_unea_income_deduction_value <> "" then
		EMWriteScreen pded_unea_income_deduction_value, 10, 62
		EMWriteScreen "X", 10, 25
		Transmit
		EMWriteScreen pded_unea_income_deduction_reason, 10, 51
		Transmit
		PF3
	End If

	'Other Earned Income Deduction
	If pded_earned_income_deduction_reason <> "" and pded_earned_income_deduction_value <> "" then
		EMWriteScreen pded_earned_income_deduction_value, 11, 62
		EMWriteScreen "X", 11, 27
		Transmit
		EMWriteScreen pded_earned_income_deduction_reason, 10, 51
		Transmit
		PF3
	End If
	
	'Extend MA-EPD Income/Asset Limits
	If pded_ma_epd_inc_asset_limit <> "" then
		pded_ma_epd_inc_asset_limit = ucase(pded_ma_epd_inc_asset_limit)
		pded_ma_epd_inc_asset_limit = left(pded_ma_epd_inc_asset_limit,1)
		EMWriteScreen pded_ma_epd_inc_asset_limit, 12, 65
	End If
	
	'====================== FIX LATER ============================
	
	'Blind/Disa Student Child Disregard
	'If  <> "" then
		'ADD LOGIC LATER
	'End If
	
	'=============================================================
	
	'Guardianship Fee
	If pded_guard_fee <> "" then
		EMWriteScreen pded_guard_fee, 15, 44
	End If
	
	'Rep Payee Fee
	If pded_rep_payee_fee <> "" then
		EMWriteScreen pded_guard_fee, 15, 70
	End If
	
	'Other Expense
	If pded_other_expense <> "" then
		EMWriteScreen pded_other_expense, 18, 41
	End If
	
	'Shelter Special Needs
	If pded_shel_spcl_needs <> "" then
		pded_shel_spcl_needs = ucase(pded_shel_spcl_needs)
		pded_shel_spcl_needs = left(pded_shel_spcl_needs,1)
		EMWriteScreen pded_shel_spcl_needs, 18, 78
	End If
	
	'Excess Need
	If pded_excess_need <> "" then
		EMWriteScreen pded_excess_need, 19, 41
	End If
	
	'Restaurant Meals
	If pded_restaurant_meals <> "" then
		pded_restaurant_meals = ucase(pded_restaurant_meals)
		pded_restaurant_meals = left(pded_restaurant_meals,1)
		EMWriteScreen pded_restaurant_meals, 19, 78
	End If
		
	Transmit
	
End Function

Function write_panel_to_maxis_HCRE(hcre_appl_addnd_date_input,hcre_retro_months_input,hcre_recvd_by_service_date_input)
	call navigate_to_screen("STAT","HCRE")
	call create_panel_if_nonexistent
	'Converting the Appl Addendum Date into a usable format
	call maxis_dater(hcre_appl_addnd_date_input, hcre_appl_addnd_date_output, "HCRE Addendum Date") 
	'Converting the Received by service date into a usable format
	call maxis_dater(hcre_recvd_by_service_date_input, hcre_recvd_by_service_date_output, "received by Service Date") 
	'Converts Retro Months Input into a negative
	hcre_retro_months_input = (Abs(hcre_retro_months_input)*(-1))
	call add_months(hcre_retro_months_input,hcre_appl_addnd_date_output,hcre_retro_date_output)
	row = 1
	col = 1
	EMSearch "* " & reference_number, row, col
		'Appl Addendum Request Date
	EMWriteScreen left(hcre_appl_addnd_date_output,2)		, row, col + 29	
	EMWriteScreen mid(hcre_recvd_by_service_date_input,4,2)	, row, col + 32	
	EMWriteScreen right(hcre_appl_addnd_date_output,2)		, row, col + 35
		'Coverage Request Date
	EMWriteScreen left(hcre_retro_date_output,2)	, row, col + 42	
	EMWriteScreen right(hcre_retro_date_output,2)	, row, col + 45
		'Recv By Sv Date
	EMWriteScreen left(hcre_recvd_by_service_date_output,2)	, row, col + 51	
	EMWriteScreen mid(hcre_recvd_by_service_date_output,4,2), row, col + 54	
	EMWriteScreen right(hcre_recvd_by_service_date_output,2), row, col + 57

	transmit
	
	'========================= REMOVE AFTER TESTING ======================
	
	call debugger()
	
	'=====================================================================
	
End Function

Function write_panel_to_maxis_INSA(insa_pers_coop_ohi,insa_good_cause_status,insa_good_cause_cliam_date,insa_good_cause_evidence,insa_coop_cost_effect,insa_insur_name,insa_prescrip_drug_cover,insa_prescrip_end_date)
	call navigate_to_screen("STAT","INSA")
	call create_panel_if_nonexistent
	
	'Resp Persons Coop With OHI
	If insa_pers_coop_ohi <> "" then
		insa_pers_coop_ohi = ucase(insa_pers_coop_ohi)
		insa_pers_coop_ohi = left(insa_pers_coop_ohi,1)
		EMWriteScreen insa_pers_coop_ohi, 19, 78
	End If
	
	'Good Cause Status
	If insa_good_cause_status <> "" then
		EMWriteScreen insa_good_cause_status, 5, 62
	End If
	
	'Good Cause Claim Date
	If insa_good_cause_cliam_date <> "" then
		call maxis_dater(insa_good_cause_cliam_date, insa_good_cause_cliam_date_output, "Good Cause Claim Date")
		EMWriteScreen left(insa_good_cause_cliam_date_output,2)	, 6, 62	
		EMWriteScreen mid(insa_good_cause_cliam_date_output,4,2), 6, 65	
		EMWriteScreen right(insa_good_cause_cliam_date_output,2), 6, 68
	End If
	
	'Good Cause Evidence
	If insa_good_cause_evidence <> "" then
		insa_good_cause_evidence = ucase(insa_good_cause_evidence)
		insa_good_cause_evidence = left(insa_good_cause_evidence,1)
		EMWriteScreen insa_good_cause_evidence, 19, 78
	End If
	
	'Coop With cost effective rqmt
	If insa_coop_cost_effect <> "" then
		insa_coop_cost_effect = ucase(insa_coop_cost_effect)
		insa_coop_cost_effect = left(insa_coop_cost_effect,1)
		EMWriteScreen insa_coop_cost_effect, 19, 78
	End If
	
	'Insurance Company
	If insa_insur_name <> "" then
		EMWriteScreen insa_insur_name, 10, 38
	End If
	
	'Prescription Drug Coverage
	If insa_prescrip_drug_cover <> "" then
		insa_prescrip_drug_cover = ucase(insa_prescrip_drug_cover)
		insa_prescrip_drug_cover = left(insa_prescrip_drug_cover,1)
		EMWriteScreen insa_prescrip_drug_cover, 11, 62
	End If
	
	'Prescription Drug Coverage End Date
	If insa_prescrip_end_date <> "" then
		call maxis_dater(insa_prescrip_end_date, insa_prescrip_end_date_output, "Good Cause Claim Date")
		EMWriteScreen left(insa_prescrip_end_date_output,2)	, 12, 62	
		EMWriteScreen mid(insa_prescrip_end_date_output,4,2), 12, 65	
		EMWriteScreen right(insa_prescrip_end_date_output,2), 12, 68
	End If
	
	'Covered Persons Ref Nbr
	row = 15
	col = 30
	insa_ref_nmb_loc = 1
	EMReadScreen insa_space_avail, 2, row, col
	Do
		If insa_space_avail <> "__" and insa_space_avail <> reference_number then
			col = col + 4
			insa_ref_nmb_loc = insa_ref_nmb_loc + 1
			If insa_ref_nmb_loc = 11 then
				row = 16
				col = 30
			ElseIf insa_ref_nmb_loc = 21 then
				msgbox "No covered person spaces are available."
			End If
		ElseIf insa_space_avail = "__" then
			EMWriteScreen reference_number, row, col
		End If
		EMReadScreen insa_space_avail, 2, row, col
	Loop until insa_space_avail = reference_number
	
	transmit

End Function

'=============== NEEDS CREATION ====================

'Function write_panel_to_maxis_DFLN()
'	call navigate_to_screen("STAT","DFLN")
'	call create_panel_if_nonexistent
'	
'End Function

'====================================================

Function write_panel_to_maxis_STWK(stwk_empl_name,stwk_wrk_stop_date,stwk_wrk_stop_date_verif,stwk_inc_stop_date,stwk_empl_yn,stwk_vol_quit,stwk_ref_empl_date,stwk_gc_cash,stwk_gc_grh,stwk_fs_pwe,stwk_maepd_ext)
	call navigate_to_screen("STAT","STWK")
	call create_panel_if_nonexistent
	
	'Employer Name
	If stwk_empl_name <> "" then
		EMWriteScreen stwk_empl_name, 6, 46
	End If 
	
	'Work Stop Date and Verif
	If stwk_wrk_stop_date <> "" then
		call maxis_dater(stwk_wrk_stop_date, stwk_wrk_stop_date_output, "Good Cause Claim Date")
		EMWriteScreen left(stwk_wrk_stop_date_output,2)	, 7, 46	
		EMWriteScreen mid(stwk_wrk_stop_date_output,4,2), 7, 49	
		EMWriteScreen right(stwk_wrk_stop_date_output,2), 7, 52
	End If
	If stwk_wrk_stop_date_verif <> "" then
		EMWriteScreen stwk_wrk_stop_date_verif, 7, 63
	End If
	
	'Income Stop Date 
	If stwk_inc_stop_date <> "" then
		call maxis_dater(stwk_inc_stop_date, stwk_inc_stop_date_output, "Good Cause Claim Date")
		EMWriteScreen left(stwk_inc_stop_date_output,2)	, 8, 46	
		EMWriteScreen mid(stwk_inc_stop_date_output,4,2), 8, 49	
		EMWriteScreen right(stwk_inc_stop_date_output,2), 8, 52
	End If
	
	'Refused Empl
	If stwk_empl_yn <> "" then
		stwk_empl_yn = ucase(stwk_empl_yn)
		stwk_empl_yn = left(stwk_empl_yn,1)
		EMWriteScreen stwk_empl_yn, 8, 78
	End If
	
	'Voluntarily Quit
	If stwk_vol_quit <> "" then
		stwk_vol_quit = ucase(stwk_vol_quit)
		stwk_vol_quit = left(stwk_vol_quit,1)
		EMWriteScreen stwk_vol_quit, 10, 46
	End If
	
	'Refused Empl Date
	If stwk_ref_empl_date <> "" then
		call maxis_dater(stwk_ref_empl_date, stwk_ref_empl_date_output, "Good Cause Claim Date")
		EMWriteScreen left(stwk_ref_empl_date_output,2)	, 10, 72	
		EMWriteScreen mid(stwk_ref_empl_date_output,4,2), 10, 75	
		EMWriteScreen right(stwk_ref_empl_date_output,2), 10, 78
	End If
	
	'Good Cause cash, grh, fs
	If stwk_gc_cash <> "" then
		stwk_gc_cash = ucase(stwk_gc_cash)
		stwk_gc_cash = left(stwk_gc_cash,1)
		EMWriteScreen stwk_gc_cash, 12, 52
	End If
	If stwk_gc_grh <> "" then
		stwk_gc_grh = ucase(stwk_gc_grh)
		stwk_gc_grh = left(stwk_gc_grh,1)
		EMWriteScreen stwk_gc_grh, 12, 60
	End If
	If stwk_gc_fs <> "" then
		stwk_gc_fs = ucase(stwk_gc_fs)
		stwk_gc_fs = left(stwk_gc_fs,1)
		EMWriteScreen stwk_gc_fs, 12, 67
	End If
	
	'FS PWE
	If stwk_fs_pwe <> "" then
		stwk_fs_pwe = ucase(stwk_fs_pwe)
		stwk_fs_pwe = left(stwk_fs_pwe,1)
		EMWriteScreen stwk_fs_pwe, 12, 67
	End If
	
	'MA-EPD Extension
	If stwk_maepd_ext <> "" then
		EMWriteScreen stwk_maepd_ext, 16, 46
	End If

	Transmit
	
End Function

Function write_panel_to_maxis_ABPS(abps_supp_coop,abps_gc_status)
	call navigate_to_screen("STAT","PARE")
	EMReadScreen abps_pare_check, 1, 2, 78
	If abps_pare_check = "0" then
		MsgBox "No PARE exists. Exiting Creating ABPS."
	ElseIf abps_pare_check <> "0" then
		child_list = ""
		row = 8
		Do
			EMReadScreen child_check, 2, row, 24
			If child_check <> "__" then
				If child_list = "" then
					child_list = child_check
				ElseIf child_list <> "" then		
					child_list = child_list & "," & child_check
				End If
			End If
			row = row + 1
			If row = 18 then
				PF8
				row = 8
			End If
		Loop until child_check = "__"
		call navigate_to_screen("STAT","ABPS")
		call create_panel_if_nonexistent		
		abps_child_list = split(child_list, ",")
		row = 15
		for each abps_child in abps_child_list
			EMWriteScreen abps_child, row, 35
			EMWriteScreen "2", row, 53
			EMWriteScreen "1", row, 67
			row = row + 1
			If row = 18 then
				PF8
				row = 15
			End If		
		next
		call maxis_dater(date(),abps_act_date,"Actual Date")
		EMWriteScreen left(abps_act_date,2)			, 18, 38
		EMWriteScreen mid(abps_act_date,4,2)		, 18, 41
		EMWriteScreen "20" & right(abps_act_date,2)	, 18, 44
		EMWriteScreen reference_number, 4, 47
		If abps_supp_coop <> "" then
			abps_supp_coop = ucase(abps_supp_coop)
			abps_supp_coop = left(abps_supp_coop,1)
			EMWriteScreen abps_supp_coop, 4, 73
		End If
		If abps_gc_status <> "" then
			EMWriteScreen abps_gc_status, 5, 47
		End If
		transmit
	End If
	
End Function

Function write_panel_to_maxis_WKEX()
	call navigate_to_screen("STAT","WKEX")
	call create_panel_if_nonexistent
	
	
	
End Function

Function write_panel_to_maxis_FMED()
	call navigate_to_screen("STAT","FMED")
	call create_panel_if_nonexistent
	
	
	
End Function

'----------- TEST AREA -------------

case_number = "211486"
reference_number = "01"

'write_panel_to_maxis_MSUR(Any_Date)
'write_panel_to_maxis_PDED(pded_wid_deduction,pded_adult_child_disregard,pded_wid_disregard,pded_unea_income_deduction_reason,pded_unea_income_deduction_value,pded_earned_income_deduction_reason,pded_earned_income_deduction_value,pded_ma_epd_inc_asset_limit,pded_guard_fee,pded_rep_payee_fee,pded_other_expense,pded_shel_spcl_needs,pded_excess_need,pded_restaurant_meals)
'write_panel_to_maxis_HCRE(hcre_appl_addnd_date_input,hcre_retro_months_input,hcre_recvd_by_service_date_input)
'write_panel_to_maxis_INSA(insa_pers_coop_ohi,insa_good_cause_status,insa_good_cause_cliam_date,insa_good_cause_evidence,insa_coop_cost_effect,insa_insur_name,insa_prescrip_drug_cover,insa_prescrip_end_date)
'write_panel_to_maxis_STWK(stwk_empl_name,stwk_wrk_stop_date,stwk_wrk_stop_date_verif,stwk_inc_stop_date,stwk_empl_yn,stwk_vol_quit,stwk_ref_empl_date,stwk_gc_cash,stwk_gc_grh,stwk_fs_pwe,stwk_maepd_ext)
'write_panel_to_maxis_ABPS(abps_supp_coop,abps_gc_status)
'
'

EMConnect ""



stopscript