'MISC functions <<<<<<<<<<MERGE INTO FUNCTIONS FILE, THANKS TO LUKE

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

'Excel Function <<<<<<<MERGE INTO FUNCTIONS FILE, THANKS TO LUKE

Function excel_open(file_url, visible_status, alerts_status, ObjExcel, objWorkbook)
	Set objExcel = CreateObject("Excel.Application") 'Allows a user to perform functions within Microsoft Excel
	objExcel.Visible = visible_status
	Set objWorkbook = objExcel.Workbooks.Open(file_url) 'Opens an excel file from a specific URL
	objExcel.DisplayAlerts = alerts_status
End Function

'BIG SLEW OF MAXIS WRITE FUNCTIONS------------------------------------------------------------------------------------------
Function write_panel_to_MAXIS_ABPS(abps_supp_coop,abps_gc_status)
	call navigate_to_screen("STAT","PARE")							'Starts by creating an array of all the kids on PARE
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
		call navigate_to_screen("STAT","ABPS")						'Navigates to ABPS to enter kids in
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
		IF abps_act_date <> "" THEN call MAXIS_dater(date, abps_act_date, "Actual Date")
		EMWriteScreen left(abps_act_date,2)			, 18, 38
		EMWriteScreen mid(abps_act_date,4,2)		, 18, 41
		EMWriteScreen "20" & right(abps_act_date,2)	, 18, 44
		EMWriteScreen reference_number, 4, 47		'Enters the reference_number
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

Function write_panel_to_MAXIS_ACCT(acct_type, acct_numb, acct_location, acct_balance, acct_bal_ver, acct_date, acct_withdraw, acct_cash_count, acct_snap_count, acct_HC_count, acct_GRH_count, acct_IV_count, acct_joint_owner, acct_share_ratio, acct_interest_date_mo, acct_interest_date_yr)
	Call Navigate_to_screen("STAT", "ACCT")  'navigates to the stat panel
	call create_panel_if_nonexistent
	Emwritescreen acct_type, 6, 44  'enters the account type code
	Emwritescreen acct_numb, 7, 44  'enters the account number
	Emwritescreen acct_location, 8, 44  'enters the account location
	Emwritescreen acct_balance, 10, 46  'enters the balance
	Emwritescreen acct_bal_ver, 10, 63  'enters the balance verification
	IF acct_date <> "" THEN call create_MAXIS_friendly_date(acct_date, 0, 11, 44)  'enters the account balance date in a MAXIS friendly format. mm/dd/yy
	Emwritescreen acct_withdraw, 12, 46  'enters the withdrawl penalty
	Emwritescreen acct_cash_count, 14, 50  'enters y/n if counted for cash
	Emwritescreen acct_snap_count, 14, 57  'enters y/n if counted for snap
	Emwritescreen acct_HC_count, 14, 64  'enters y/n if counted for HC
	Emwritescreen acct_GRH_count, 14, 72  'enters y/n if counted for grh
	Emwritescreen acct_IV_count, 14, 80  'enters y/n if counted for IV
	Emwritescreen acct_joint_owner, 15, 44  'enters if it is a jointly owned acct
	Emwritescreen left(acct_share_ratio, 1), 15, 76  'enters the ratio of ownership using the left 1 digit of what is entered into the file
	Emwritescreen right(acct_share_ratio, 1), 15, 80  'enters the ratio of ownership using the right 1 digit of what is entered into the file
	Emwritescreen acct_interest_date_mo, 17, 57  'enters the next interest date MM format
	Emwritescreen acct_interest_date_yr, 17, 60  'enters the next interest date YY format
	transmit
	transmit
End Function

FUNCTION write_panel_to_MAXIS_ACUT(ACUT_shared, ACUT_heat, ACUT_air, ACUT_electric, ACUT_fuel, ACUT_garbage, ACUT_water, ACUT_sewer, ACUT_other, ACUT_phone, ACUT_heat_verif, ACUT_air_verif, ACUT_electric_verif, ACUT_fuel_verif, ACUT_garbage_verif, ACUT_water_verif, ACUT_sewer_verif, ACUT_other_verif)
	call navigate_to_screen("STAT", "ACUT")
	call create_panel_if_nonexistent
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

'---This function writes the information for BILS.
FUNCTION write_panel_to_MAXIS_BILS(bils_1_ref_num, bils_1_serv_date, bils_1_serv_type, bils_1_gross_amt, bils_1_third_party, bils_1_verif, bils_1_bils_type, bils_2_ref_num, bils_2_serv_date, bils_2_serv_type, bils_2_gross_amt, bils_2_third_party, bils_2_verif, bils_2_bils_type, bils_3_ref_num, bils_3_serv_date, bils_3_serv_type, bils_3_gross_amt, bils_3_third_party, bils_3_verif, bils_3_bils_type, bils_4_ref_num, bils_4_serv_date, bils_4_serv_type, bils_4_gross_amt, bils_4_third_party, bils_4_verif, bils_4_bils_type, bils_5_ref_num, bils_5_serv_date, bils_5_serv_type, bils_5_gross_amt, bils_5_third_party, bils_5_verif, bils_5_bils_type, bils_6_ref_num, bils_6_serv_date, bils_6_serv_type, bils_6_gross_amt, bils_6_third_party, bils_6_verif, bils_6_bils_type, bils_7_ref_num, bils_7_serv_date, bils_7_serv_type, bils_7_gross_amt, bils_7_third_party, bils_7_verif, bils_7_bils_type, bils_8_ref_num, bils_8_serv_date, bils_8_serv_type, bils_8_gross_amt, bils_8_third_party, bils_8_verif, bils_8_bils_type, bils_9_ref_num, bils_9_serv_date, bils_9_serv_type, bils_9_gross_amt, bils_9_third_party, bils_9_verif, bils_9_bils_type)
	CALL navigate_to_screen("STAT", "BILS")
	ERRR_screen_check
	EMReadScreen num_of_BILS, 1, 2, 78
	IF num_of_BILS = "0" THEN
		EMWriteScreen "NN", 20, 79
		transmit
	ELSE
		PF9
	END IF
	
	'---MAXIS will not allow BILS to be updated if HC is inactive. Exiting the function if HC is inactive.
	EMReadScreen hc_inactive, 21, 24, 2
	IF hc_inactive = "HC STATUS IS INACTIVE" THEN Exit FUNCTION
	
	BILS_row = 6
	DO
		EMReadScreen available_row, 2, BILS_row, 26
		IF available_row <> "__" THEN BILS_row = BILS_row + 1
		IF BILS_row = 18 THEN 
			PF20
			BILS_row = 6
		END IF
	LOOP UNTIL available_row = "__"
	
	IF bils_1_ref_num <> "" THEN 
		IF len(bils_1_ref_num) = 1 THEN bils_1_ref_num = "0" & bils_1_ref_num
		EMWriteScreen bils_1_ref_num, BILS_row, 26
		CALL create_MAXIS_friendly_date(bils_1_serv_date, 0, BILS_row, 30)
		EMWriteScreen bils_1_serv_type, BILS_row, 40
		EMWriteScreen bils_1_gross_amt, BILS_row, 45
		EMWriteScreen bils_1_third_party, BILS_row, 57
		IF bils_1_verif = "03" AND bils_1_serv_type <> "22" THEN bils_1_verif = "06"	'Because CL Stmt is only acceptable for Medical transportation
		EMWriteScreen bils_1_verif, BILS_row, 67
		EMWriteScreen bils_1_bils_type, BILS_row, 71
		BILS_row = BILS_row + 1
		IF BILS_row = 18 THEN 
			PF20
			BILS_row = 6
		END IF
	END IF
	IF bils_2_ref_num <> "" THEN 
		IF len(bils_2_ref_num) = 1 THEN bils_2_ref_num = "0" & bils_2_ref_num
		EMWriteScreen bils_2_ref_num, BILS_row, 26
		CALL create_MAXIS_friendly_date(bils_2_serv_date, 0, BILS_row, 30)
		EMWriteScreen bils_2_serv_type, BILS_row, 40
		EMWriteScreen bils_2_gross_amt, BILS_row, 45
		EMWriteScreen bils_2_third_party, BILS_row, 57
		IF bils_2_verif = "03" AND bils_2_serv_type <> "22" THEN bils_2_verif = "06"	'Because CL Stmt is only acceptable for Medical transportation
		EMWriteScreen bils_2_verif, BILS_row, 67
		EMWriteScreen bils_2_bils_type, BILS_row, 71
		BILS_row = BILS_row + 1
		IF BILS_row = 18 THEN 
			PF20
			BILS_row = 6
		END IF
	END IF
	IF bils_3_ref_num <> "" THEN 
		IF len(bils_3_ref_num) = 1 THEN bils_3_ref_num = "0" & bils_3_ref_num
		EMWriteScreen bils_3_ref_num, BILS_row, 26
		CALL create_MAXIS_friendly_date(bils_3_serv_date, 0, BILS_row, 30)
		EMWriteScreen bils_3_serv_type, BILS_row, 40
		EMWriteScreen bils_3_gross_amt, BILS_row, 45
		EMWriteScreen bils_3_third_party, BILS_row, 57
		IF bils_3_verif = "03" AND bils_3_serv_type <> "22" THEN bils_3_verif = "06"	'Because CL Stmt is only acceptable for Medical transportation
		EMWriteScreen bils_3_verif, BILS_row, 67
		EMWriteScreen bils_3_bils_type, BILS_row, 71
		BILS_row = BILS_row + 1
		IF BILS_row = 18 THEN 
			PF20
			BILS_row = 6
		END IF
	END IF
	IF bils_4_ref_num <> "" THEN
		IF len(bils_4_ref_num) = 1 THEN bils_4_ref_num = "0" & bils_4_ref_num
		EMWriteScreen bils_4_ref_num, BILS_row, 26
		CALL create_MAXIS_friendly_date(bils_4_serv_date, 0, BILS_row, 30)
		EMWriteScreen bils_4_serv_type, BILS_row, 40
		EMWriteScreen bils_4_gross_amt, BILS_row, 45
		EMWriteScreen bils_4_third_party, BILS_row, 57
		IF bils_4_verif = "03" AND bils_4_serv_type <> "22" THEN bils_4_verif = "06"	'Because CL Stmt is only acceptable for Medical transportation
		EMWriteScreen bils_4_verif, BILS_row, 67
		EMWriteScreen bils_4_bils_type, BILS_row, 71
		BILS_row = BILS_row + 1
		IF BILS_row = 18 THEN 
			PF20
			BILS_row = 6
		END IF
	END IF
	IF bils_5_ref_num <> "" THEN 
		IF len(bils_5_ref_num) = 1 THEN bils_5_ref_num = "0" & bils_5_ref_num
		EMWriteScreen bils_5_ref_num, BILS_row, 26
		CALL create_MAXIS_friendly_date(bils_5_serv_date, 0, BILS_row, 30)
		EMWriteScreen bils_5_serv_type, BILS_row, 40
		EMWriteScreen bils_5_gross_amt, BILS_row, 45
		EMWriteScreen bils_5_third_party, BILS_row, 57
		IF bils_5_verif = "03" AND bils_5_serv_type <> "22" THEN bils_5_verif = "06"	'Because CL Stmt is only acceptable for Medical transportation
		EMWriteScreen bils_5_verif, BILS_row, 67
		EMWriteScreen bils_5_bils_type, BILS_row, 71
		BILS_row = BILS_row + 1
		IF BILS_row = 18 THEN 
			PF20
			BILS_row = 6
		END IF
	END IF
	IF bils_6_ref_num <> "" THEN 
		IF len(bils_6_ref_num) = 1 THEN bils_6_ref_num = "0" & bils_6_ref_num
		EMWriteScreen bils_6_ref_num, BILS_row, 26
		CALL create_MAXIS_friendly_date(bils_6_serv_date, 0, BILS_row, 30)
		EMWriteScreen bils_6_serv_type, BILS_row, 40
		EMWriteScreen bils_6_gross_amt, BILS_row, 45
		EMWriteScreen bils_6_third_party, BILS_row, 57
		IF bils_6_verif = "03" AND bils_6_serv_type <> "22" THEN bils_6_verif = "06"	'Because CL Stmt is only acceptable for Medical transportation
		EMWriteScreen bils_6_verif, BILS_row, 67
		EMWriteScreen bils_6_bils_type, BILS_row, 71
		BILS_row = BILS_row + 1
		IF BILS_row = 18 THEN 
			PF20
			BILS_row = 6
		END IF
	END IF
	IF bils_7_ref_num <> "" THEN 
		IF len(bils_7_ref_num) = 1 THEN bils_7_ref_num = "0" & bils_7_ref_num
		EMWriteScreen bils_7_ref_num, BILS_row, 26
		CALL create_MAXIS_friendly_date(bils_7_serv_date, 0, BILS_row, 30)
		EMWriteScreen bils_7_serv_type, BILS_row, 40
		EMWriteScreen bils_7_gross_amt, BILS_row, 45
		EMWriteScreen bils_7_third_party, BILS_row, 57
		IF bils_7_verif = "03" AND bils_7_serv_type <> "22" THEN bils_7_verif = "06"	'Because CL Stmt is only acceptable for Medical transportation
		EMWriteScreen bils_7_verif, BILS_row, 67
		EMWriteScreen bils_7_bils_type, BILS_row, 71
		BILS_row = BILS_row + 1
		IF BILS_row = 18 THEN 
			PF20
			BILS_row = 6
		END IF
	END IF
	IF bils_8_ref_num <> "" THEN 
		IF len(bils_8_ref_num) = 1 THEN bils_8_ref_num = "0" & bils_8_ref_num
		EMWriteScreen bils_8_ref_num, BILS_row, 26
		CALL create_MAXIS_friendly_date(bils_8_serv_date, 0, BILS_row, 30)
		EMWriteScreen bils_8_serv_type, BILS_row, 40
		EMWriteScreen bils_8_gross_amt, BILS_row, 45
		EMWriteScreen bils_8_third_party, BILS_row, 57
		IF bils_8_verif = "03" AND bils_8_serv_type <> "22" THEN bils_8_verif = "06"	'Because CL Stmt is only acceptable for Medical transportation
		EMWriteScreen bils_8_verif, BILS_row, 67
		EMWriteScreen bils_8_bils_type, BILS_row, 71
		BILS_row = BILS_row + 1
		IF BILS_row = 18 THEN 
			PF20
			BILS_row = 6
		END IF
	END IF
	IF bils_9_ref_num <> "" THEN 
		IF len(bils_9_ref_num) = 1 THEN bils_9_ref_num = "0" & bils_9_ref_num
		EMWriteScreen bils_9_ref_num, BILS_row, 26
		CALL create_MAXIS_friendly_date(bils_9_serv_date, 0, BILS_row, 30)
		EMWriteScreen bils_9_serv_type, BILS_row, 40
		EMWriteScreen bils_9_gross_amt, BILS_row, 45
		EMWriteScreen bils_9_third_party, BILS_row, 57
		IF bils_9_verif = "03" AND bils_9_serv_type <> "22" THEN bils_9_verif = "06"	'Because CL Stmt is only acceptable for Medical transportation
		EMWriteScreen bils_9_verif, BILS_row, 67
		EMWriteScreen bils_9_bils_type, BILS_row, 71
	END IF
END FUNCTION


'---This function writes using the variables read off of the specialized excel template to the busi panel in MAXIS
Function write_panel_to_MAXIS_BUSI(busi_type, busi_start_date, busi_end_date, busi_cash_total_retro, busi_cash_total_prosp, busi_cash_total_ver, busi_IV_total_prosp, busi_IV_total_ver, busi_snap_total_retro, busi_snap_total_prosp, busi_snap_total_ver, busi_hc_total_prosp_a, busi_hc_total_ver_a, busi_hc_total_prosp_b, busi_hc_total_ver_b, busi_cash_exp_retro, busi_cash_exp_prosp, busi_cash_exp_ver, busi_IV_exp_prosp, busi_IV_exp_ver, busi_snap_exp_retro, busi_snap_exp_prosp, busi_snap_exp_ver, busi_hc_exp_prosp_a, busi_hc_exp_ver_a, busi_hc_exp_prosp_b, busi_hc_exp_ver_b, busi_retro_hours, busi_prosp_hours, busi_hc_total_est_a, busi_hc_total_est_b, busi_hc_exp_est_a, busi_hc_exp_est_b, busi_hc_hours_est)
	Call navigate_to_screen("STAT", "BUSI")  'navigates to the stat panel
	Emwritescreen reference_number, 20, 76
	transmit
	
	EMReadScreen num_of_BUSI, 1, 2, 78
	IF num_of_BUSI = "0" THEN 
		EMWriteScreen "__", 20, 76
		EMWriteScreen "NN", 20, 79
		transmit
		
		'Reading the footer month, converting to an actual date, we'll need this for determining if the panel is 02/15 or later (there was a change in 02/15 which moved stuff)
		EMReadScreen BUSI_footer_month, 5, 20, 55
		BUSI_footer_month = replace(BUSI_footer_month, " ", "/01/")
		'Treats panels older than 02/15 with the old logic, because the panel was changed in 02/15. We'll remove this in August if all goes well.
		If datediff("d", "02/01/2015", BUSI_footer_month) < 0 then 
			Emwritescreen busi_type, 5, 37  'enters self employment type
			call create_MAXIS_friendly_date(busi_start_date, 0, 5, 54)  'enters self employment start date in MAXIS friendly format mm/dd/yy
			IF busi_end_date <> "" THEN call create_MAXIS_friendly_date(busi_end_date, 0, 5, 71)  'enters self employment start date in MAXIS friendly format mm/dd/yy
			Emwritescreen "x", 7, 26  'this enters into the gross income calculator
			Transmit
			Do
				Emreadscreen busi_gross_income_check, 12, 06, 35  'This checks to see if the gross income calculator has actually opened. 
			LOOP UNTIL busi_gross_income_check = "Gross Income"
			Emwritescreen busi_cash_total_retro, 9, 43  'enters the cash total income retrospective number
			Emwritescreen busi_cash_total_prosp, 9, 59  'enters the cash total income prospective number
			Emwritescreen busi_cash_total_ver, 9, 73    'enters the cash total income verification code
			Emwritescreen busi_IV_total_prosp, 10, 59   'enters the IV total income prospective number
			Emwritescreen busi_IV_total_ver, 10, 73     'enters the IV total income verification code
			Emwritescreen busi_snap_total_retro, 11, 43 'enters the snap total income retro number
			Emwritescreen busi_snap_total_prosp, 11, 59 'enters the snap total income prosp number
			Emwritescreen busi_snap_total_ver, 11, 73   'enters the snap total verification code
			Emwritescreen busi_hc_total_prosp_a, 12, 59 'enters the HC total income prospective number for method a
			Emwritescreen busi_hc_total_ver_a, 12, 73   'enters the HC total income verification code for method a
			Emwritescreen busi_hc_total_prosp_b, 13, 59 'enters the HC total income prospective number for method b
			Emwritescreen busi_hc_total_ver_b, 13, 73   'enters the HC total income verification code for method b
			Emwritescreen busi_cash_exp_retro, 15, 43   'enters the cash expenses retrospective number
			Emwritescreen busi_cash_exp_prosp, 15, 59   'enters the cash expenses prospective number
			Emwritescreen busi_cash_exp_ver, 15, 73     'enters the cash expenses verification code
			Emwritescreen busi_IV_exp_prosp, 16, 59     'enters the IV expenses retro number
			Emwritescreen busi_IV_exp_ver, 9, 73        'enters the IV expenses verification code
			Emwritescreen busi_snap_exp_retro, 17, 43   'enters the snap expenses retro number
			Emwritescreen busi_snap_exp_prosp, 17, 59   'enters the snap expenses prospective number
			Emwritescreen busi_snap_exp_ver, 17, 73     'enters the snap expenses verif code
			Emwritescreen busi_hc_exp_prosp_a, 18, 59   'enters the hc expenses prospective number for method a
			Emwritescreen busi_hc_exp_ver_a, 18, 73    'enters the hc expenses verification code for method a
			Emwritescreen busi_hc_exp_prosp_b, 19, 59   'enters the hc expenses prospective number for method b
			Emwritescreen busi_hc_exp_ver_b, 19, 73	  'enters the hc expenses verification code for method b
			transmit
			PF3
			Emwritescreen busi_retro_hours, 14, 59  'enters the retrospective hours
			Emwritescreen busi_prosp_hours, 14, 73  'enters the prospective hours
		
		ELSE				'This is the NEW logic for all months after 02/2015
			Emwritescreen busi_type, 5, 37  'enters self employment type
			call create_MAXIS_friendly_date(busi_start_date, 0, 5, 55)  'enters self employment start date in MAXIS friendly format mm/dd/yy
			IF busi_end_date <> "" THEN call create_MAXIS_friendly_date(busi_end_date, 0, 5, 72)  'enters self employment start date in MAXIS friendly format mm/dd/yy
			Emwritescreen "x", 6, 26  'this enters into the gross income calculator
			Transmit		
			Do
				Emreadscreen busi_gross_income_check, 12, 06, 35  'This checks to see if the gross income calculator has actually opened. 
			LOOP UNTIL busi_gross_income_check = "Gross Income"		
			Emwritescreen busi_cash_total_retro, 9, 43  'enters the cash total income retrospective number
			Emwritescreen busi_cash_total_prosp, 9, 59  'enters the cash total income prospective number
			Emwritescreen busi_cash_total_ver, 9, 73    'enters the cash total income verification code
			Emwritescreen busi_IV_total_prosp, 10, 59   'enters the IV total income prospective number
			Emwritescreen busi_IV_total_ver, 10, 73     'enters the IV total income verification code
			Emwritescreen busi_snap_total_retro, 11, 43 'enters the snap total income retro number
			Emwritescreen busi_snap_total_prosp, 11, 59 'enters the snap total income prosp number
			Emwritescreen busi_snap_total_ver, 11, 73   'enters the snap total verification code
			Emwritescreen busi_hc_total_prosp_a, 12, 59 'enters the HC total income prospective number for method a
			Emwritescreen busi_hc_total_ver_a, 12, 73   'enters the HC total income verification code for method a
			Emwritescreen busi_hc_total_prosp_b, 13, 59 'enters the HC total income prospective number for method b
			Emwritescreen busi_hc_total_ver_b, 13, 73   'enters the HC total income verification code for method b
			Emwritescreen busi_cash_exp_retro, 15, 43   'enters the cash expenses retrospective number
			Emwritescreen busi_cash_exp_prosp, 15, 59   'enters the cash expenses prospective number
			Emwritescreen busi_cash_exp_ver, 15, 73     'enters the cash expenses verification code
			Emwritescreen busi_IV_exp_prosp, 16, 59     'enters the IV expenses retro number
			Emwritescreen busi_IV_exp_ver, 9, 73        'enters the IV expenses verification code
			Emwritescreen busi_snap_exp_retro, 17, 43   'enters the snap expenses retro number
			Emwritescreen busi_snap_exp_prosp, 17, 59   'enters the snap expenses prospective number
			Emwritescreen busi_snap_exp_ver, 17, 73     'enters the snap expenses verif code
			Emwritescreen busi_hc_exp_prosp_a, 18, 59   'enters the hc expenses prospective number for method a
			Emwritescreen busi_hc_exp_ver_a, 18, 73    'enters the hc expenses verification code for method a
			Emwritescreen busi_hc_exp_prosp_b, 19, 59   'enters the hc expenses prospective number for method b
			Emwritescreen busi_hc_exp_ver_b, 19, 73	  'enters the hc expenses verification code for method b
			transmit
			PF3
			Emwritescreen busi_retro_hours, 13, 60  'enters the retrospective hours
			Emwritescreen busi_prosp_hours, 13, 74  'enters the prospective hours
			'---Adding Self-Employment Method -- Hard-Coded for now.
			EMWriteScreen "01", 16, 53
			CALL create_MAXIS_friendly_date(#02/01/2015#, 0, 16, 63)
		END IF
	ELSEIF num_of_BUSI <> "0" THEN
		PF9
		'Reading the footer month, converting to an actual date, we'll need this for determining if the panel is 02/15 or later (there was a change in 02/15 which moved stuff)
		EMReadScreen BUSI_footer_month, 5, 20, 55
		BUSI_footer_month = replace(BUSI_footer_month, " ", "/01/")
		'Treats panels older than 02/15 with the old logic, because the panel was changed in 02/15. We'll remove this in August if all goes well.
		If datediff("d", "02/01/2015", BUSI_footer_month) >= 0 then 
			'---Adding Self-Employment Method -- Hard-Coded for now.
			EMWriteScreen "01", 16, 53
			CALL create_MAXIS_friendly_date(#02/01/2015#, 0, 16, 63)
			'---Going into the HC Income Estimate
			EMWriteScreen "X", 17, 27
			transmit
			DO
				EMReadScreen hc_income, 9, 4, 42
			LOOP UNTIL hc_income = "HC Income"
			EMReadScreen current_month_plus_one, 17, 21, 59
			IF current_month_plus_one = "CURRENT MONTH + 1" THEN 
				PF3
			ELSE
				Emwritescreen busi_hc_total_est_a, 7, 54                'enters hc total income estimation for method A
				Emwritescreen busi_hc_total_est_b, 8, 54                'enters hc total income estimation for method B
				Emwritescreen busi_hc_exp_est_a, 11, 54                 'enters hc expense estimation for method A
				Emwritescreen busi_hc_exp_est_b, 12, 54                 'enters hc expense estimation for method B
				Emwritescreen busi_hc_hours_est, 18, 58                 'enters hc hours estimation
				transmit
				PF3
			END IF
		END IF
	END IF
end function

Function write_panel_to_MAXIS_CARS(cars_type, cars_year, cars_make, cars_model, cars_trade_in, cars_loan, cars_value_source, cars_ownership_ver, cars_amount_owed, cars_amount_owed_ver, cars_date, cars_use, cars_HC_benefit, cars_joint_owner, cars_share_ratio)
	Call Navigate_to_screen("STAT", "CARS")  'navigates to the stat screen
	call create_panel_if_nonexistent
	Emwritescreen cars_type, 6, 43  'enters the vehicle type
	Emwritescreen cars_year, 8, 31  'enters the vehicle year
	Emwritescreen cars_make, 8, 43  'enters the vehicle make
	Emwritescreen cars_model, 8, 66  'enters the vehicle model
	Emwritescreen cars_trade_in, 9, 45  'enters the trade in value
	Emwritescreen cars_loan, 9, 62  'enters the loan value
	Emwritescreen cars_value_source, 9, 80  'enters the source of value information
	Emwritescreen cars_ownership_ver, 10, 60  'enters the ownership verification code
	Emwritescreen cars_amount_owed, 12, 45  'enters the amount owed on vehicle
	Emwritescreen cars_amount_owed_ver, 12, 60  'enters the amount owed verification code
	IF cars_date <> "" THEN call create_MAXIS_friendly_date(cars_date, 0, 13, 43)  'enters the amounted owed as of date in a MAXIS friendly format. mm/dd/yy
	Emwritescreen cars_use, 15, 43  'enters the use code for the vehicle
	Emwritescreen cars_HC_benefit, 15, 76  'enters if the vehicle is for client benefit
	Emwritescreen cars_joint_owner, 16, 43  'enters if it is a jointly owned car
	Emwritescreen left(cars_share_ratio, 1), 16, 76  'enters the ratio of ownership using the left 1 digit of what is entered into the file
	Emwritescreen right(cars_share_ratio, 1), 16, 80  'enters the ratio of ownership using the right 1 digit of what is entered into the file
End Function

'---This function writes using the variables read off of the specialized excel template to the cash panel in MAXIS
Function write_panel_to_MAXIS_CASH(cash_amount)
	Call navigate_to_screen("STAT", "CASH")  'navigates to the stat panel
	call create_panel_if_nonexistent
	Emwritescreen cash_amount, 8, 39
End Function

'---This function writes using the variables read off of the specialized excel template to the COEX panel in MAXIS.
FUNCTION write_panel_to_MAXIS_COEX(retro_support, prosp_support, support_verif, retro_alimony, prosp_alimony, alimony_verif, retro_tax_dep, prosp_tax_dep, tax_dep_verif, retro_other, prosp_other, other_verif, change_in_circum, hc_exp_support, hc_exp_alimony, hc_exp_tax_dep, hc_exp_other)
	CALL navigate_to_MAXIS_screen("STAT", "COEX")
	ERRR_screen_check
	EMWriteScreen reference_number, 20, 76
	transmit
	
	EMReadScreen num_of_COEX, 1, 2, 78
	IF num_of_COEX = "0" THEN 
		EMWriteScreen "__", 20, 76
		EMWriteScreen "NN", 20, 79
		transmit
		'---If the script is creating a new COEX panel, it will enter this information...
		EMWriteScreen support_verif, 10, 36
		EMWriteScreen retro_support, 10, 45
		EMWriteScreen prosp_support, 10, 63
		EMWriteScreen alimony_verif, 11, 36
		EMWriteScreen retro_alimony, 11, 45
		EMWriteScreen prosp_alimony, 11, 63
		EMWriteScreen tax_dep_verif, 12, 36
		EMWriteScreen retro_tax_dep, 12, 45
		EMWriteScreen prosp_tax_dep, 12, 63
		EMWriteScreen other_verif, 13, 36
		EMWriteScreen retro_other, 13, 45
		EMWriteScreen prosp_other, 13, 63
		EMWriteScreen change_in_circum, 17, 61
	ELSEIF num_of_COEX <> "0" THEN
		PF9
		'---...if the script is PF9'ing, it is doing so to enter information into the HC Expense sub-menu
		'Opening the HC Expenses Sub-menu
		EMWriteScreen "X", 18, 44
		transmit
			
		DO
			EMReadScreen hc_expense_est, 14, 4, 30
		LOOP UNTIL hc_expense_est = "HC Expense Est"
		
		EMReadScreen current_month_plus_one, 17, 13, 51
		IF current_month_plus_one <> "CURRENT MONTH + 1" THEN
			EMWriteScreen hc_exp_support, 6, 38
			EMWriteScreen hc_exp_alimony, 7, 38
			EMWriteScreen hc_exp_tax_dep, 8, 38
			EMWriteScreen hc_exp_other, 9, 38
			transmit
		END IF
		PF3
	END IF
	transmit
END FUNCTION


FUNCTION write_panel_to_MAXIS_DCEX(DCEX_provider, DCEX_reason, DCEX_subsidy, DCEX_child_number1, DCEX_child_number1_ver, DCEX_child_number1_retro, DCEX_child_number1_pro, DCEX_child_number2, DCEX_child_number2_ver, DCEX_child_number2_retro, DCEX_child_number2_pro, DCEX_child_number3, DCEX_child_number3_ver, DCEX_child_number3_retro, DCEX_child_number3_pro, DCEX_child_number4, DCEX_child_number4_ver, DCEX_child_number4_retro, DCEX_child_number4_pro, DCEX_child_number5, DCEX_child_number5_ver, DCEX_child_number5_retro, DCEX_child_number5_pro, DCEX_child_number6, DCEX_child_number6_ver, DCEX_child_number6_retro, DCEX_child_number6_pro)
	call navigate_to_screen("STAT", "DCEX") 
	EMWriteScreen reference_number, 20, 76
	transmit
	
	EMReadScreen num_of_DCEX, 1, 2, 78
	IF num_of_DCEX = "0" THEN 
		EMWriteScreen "__", 20, 76
		Emwritescreen "NN", 20, 79
		transmit
		
		'---If the script is creating a new DCEX panel, it is going to enter this information into the DCEX main screen...
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
	ELSE
		PF9
		'---...if the script is PF9'ing, it is ONLY because it is going to enter information in the HC Expense sub-menu.
		'---Writing in the HC Expenses Est
		EMWriteScreen "X", 17, 55
		transmit
		
		DO			'---Waiting to make sure the HC Expense Est window has opened.
			EMReadScreen hc_expense_est, 10, 4, 41
		LOOP UNTIL hc_expense_est = "HC Expense"
			
		EMReadScreen hc_month, 17, 18, 62
		IF hc_month = "CURRENT MONTH + 1" THEN
			PF3
		ELSE
			EMWritescreen DCEX_child_number1, 8, 39
			EMWritescreen DCEX_child_number2, 9, 39
			EMWritescreen DCEX_child_number3, 10, 39
			EMWritescreen DCEX_child_number4, 11, 39
			EMWritescreen DCEX_child_number5, 12, 39
			EMWritescreen DCEX_child_number6, 13, 39
			EMWritescreen DCEX_child_number1_pro, 8, 49
			EMWritescreen DCEX_child_number2_pro, 9, 49
			EMWritescreen DCEX_child_number3_pro, 10, 49
			EMWritescreen DCEX_child_number4_pro, 11, 49
			EMWritescreen DCEX_child_number5_pro, 12, 49
			EMWritescreen DCEX_child_number6_pro, 13, 49
			transmit
			PF3
		END IF
	END IF	
	transmit
End function

FUNCTION write_panel_to_MAXIS_DFLN(conv_dt_1, conv_juris_1, conv_st_1, conv_dt_2, conv_juris_2, conv_st_2, rnd_test_dt_1, rnd_test_provider_1, rnd_test_result_1, rnd_test_dt_2, rnd_test_provider_2, rnd_test_result_2)
	CALL navigate_to_screen("STAT", "DFLN")
	EMReadScreen num_of_DFLN, 1, 2, 78
	IF num_of_DFLN = "0" THEN
		EMWriteScreen reference_number, 20, 76
		EMWriteScreen "NN", 20, 79
		transmit
	ELSE
		PF9
	END IF
	
	CALL create_MAXIS_friendly_date(conv_dt_1, 0, 6, 27)
	EMWriteScreen conv_juris_1, 6, 41
	EMWriteScreen conv_st_1, 6, 75
	IF conv_dt_2 <> "" THEN 
		CALL create_MAXIS_friendly_date(conv_dt_2, 0, 7, 27)
		EMWriteScreen conv_juris_2, 7, 41
		EMWriteScreen conv_st_2, 7, 75
	END IF
	IF rnd_test_dt_1 <> "" THEN 
		CALL create_MAXIS_friendly_date(rnd_test_dt_1, 0, 14, 27)
		EMWriteScreen rnd_test_provider_1, 14, 41
		EMWriteScreen rnd_test_result_1, 14, 75
		IF rnd_test_dt_2 <> "" THEN 
			CALL create_MAXIS_friendly_date(rnd_test_dt_2, 0, 15, 27)
			EMWriteScreen rnd_test_provider_2, 15, 41
			EMWriteScreen rnd_test_result_2, 15, 75
		END IF
	END IF
END FUNCTION

FUNCTION write_panel_to_MAXIS_DIET(DIET_mfip_1, DIET_mfip_1_ver, DIET_mfip_2, DIET_mfip_2_ver, DIET_msa_1, DIET_msa_1_ver, DIET_msa_2, DIET_msa_2_ver, DIET_msa_3, DIET_msa_3_ver, DIET_msa_4, DIET_msa_4_ver)
	call navigate_to_screen("STAT", "DIET")
	EMWriteScreen reference_number, 20, 76
	EMWriteScreen "NN", 20, 79
	transmit

	EMWriteScreen DIET_mfip_1, 8, 40
	EMWriteScreen DIET_mfip_1_ver, 8, 51
	EMWriteScreen DIET_mfip_2, 9, 40
	EMWriteScreen DIET_mfip_2_ver, 9, 51
	EMWriteScreen DIET_msa_1, 11, 40
	EMWriteScreen DIET_msa_1_ver, 11, 51
	EMWriteScreen DIET_msa_2, 12, 40
	EMWriteScreen DIET_msa_2_ver, 12, 51
	EMWriteScreen DIET_msa_3, 13, 40
	EMWriteScreen DIET_msa_3_ver, 13, 51
	EMWriteScreen DIET_msa_4, 14, 40
	EMWriteScreen DIET_msa_4_ver, 14, 51
	transmit
END FUNCTION

'---This function writes using the variables read off of the specialized excel template to the disa panel in MAXIS
Function write_panel_to_MAXIS_DISA(disa_begin_date, disa_end_date, disa_cert_begin, disa_cert_end, disa_wavr_begin, disa_wavr_end, disa_grh_begin, disa_grh_end, disa_cash_status, disa_cash_status_ver, disa_snap_status, disa_snap_status_ver, disa_hc_status, disa_hc_status_ver, disa_waiver, disa_drug_alcohol)
	Call navigate_to_screen("STAT", "DISA")  'navigates to the stat panel
	call create_panel_if_nonexistent
	IF disa_begin_date <> "" THEN 
		call create_MAXIS_friendly_date(disa_begin_date, 0, 6, 47)  'enters the disability begin date in a MAXIS friendly format. mm/dd/yy
		EMWriteScreen DatePart("YYYY", disa_begin_date), 6, 53
	END IF
	IF disa_end_date <> "" THEN 
		call create_MAXIS_friendly_date(disa_end_date, 0, 6, 69)  'enters the disability end date in a MAXIS friendly format. mm/dd/yy
		EMWriteScreen DatePart("YYYY", disa_end_date), 6, 75
	END IF
	IF disa_cert_begin <> "" THEN
		call create_MAXIS_friendly_date(disa_cert_begin, 0, 7, 47)  'enters the disability certification begin date in a MAXIS friendly format. mm/dd/yy
		EMWriteScreen DatePart("YYYY", disa_cert_begin), 7, 53
	END IF
	IF disa_cert_end <> "" THEN
		call create_MAXIS_friendly_date(disa_cert_end, 0, 7, 69)  'enters the disability certification end date in a MAXIS friendly format. mm/dd/yy
		EMWriteScreen DatePart("YYYY", disa_cert_end), 7, 75
	END IF
	IF disa_wavr_begin <> "" THEN
		call create_MAXIS_friendly_date(disa_wavr_begin, 0, 8, 47)  'enters the disability waiver begin date in a MAXIS friendly format. mm/dd/yy
		EMWriteScreen DatePart("YYYY", disa_wavr_begin), 8, 53
	END IF
	IF disa_wavr_end <> "" THEN 
		call create_MAXIS_friendly_date(disa_wavr_end, 0, 8, 69)  'enters the disability waiver end date in a MAXIS friendly format. mm/dd/yy
		EMWriteScreen DatePart("YYYY", disa_wavr_end), 8, 75
	END IF
	IF disa_grh_begin <> "" THEN 
		call create_MAXIS_friendly_date(disa_grh_begin, 0, 9, 47)  'enters the disability grh begin date in a MAXIS friendly format. mm/dd/yy
		EMWriteScreen DatePart("YYYY", disa_grh_begin), 9, 53
	END IF
	IF disa_grh_end <> "" THEN 
		call create_MAXIS_friendly_date(disa_grh_end, 0, 9, 69)  'enters the disability grh end date in a MAXIS friendly format. mm/dd/yy
		EMWriteScreen DatePart("YYYY", disa_grh_end), 9, 75
	END IF
	Emwritescreen disa_cash_status, 11, 59  'enters status code for cash disa status
	Emwritescreen disa_cash_status_ver, 11, 69  'enters verification code for cash disa status
	Emwritescreen disa_snap_status, 12, 59  'enters status code for snap disa status
	Emwritescreen disa_snap_status_ver, 12, 69  'enters verification code for snap disa status
	Emwritescreen disa_hc_status, 13, 59  'enters status code for hc disa status
	Emwritescreen disa_hc_status_ver, 13, 69  'enters verification code for hc disa status
	Emwritescreen disa_waiver, 14, 59  'enters home and comminuty waiver code
	Emwritescreen disa_1619, 16, 59  'enters 1619 status
	Emwritescreen disa_drug_alcohol, 18, 69  'enters material drug & alcohol verification
End Function

Function write_panel_to_MAXIS_DSTT(DSTT_ongoing_income, DSTT_HH_income_stop_date, DSTT_income_expected_amt)
	call navigate_to_screen("STAT", "DSTT")
	call ERRR_screen_check
	call create_panel_if_nonexistent
	EMWriteScreen DSTT_ongoing_income, 6, 69
	IF HH_income_stop_date <> "" THEN call create_MAXIS_friendly_date(HH_income_stop_date, 0, 9, 69)
	EMWriteScreen income_expected_amt, 12, 71
End function

FUNCTION write_panel_to_MAXIS_EATS(eats_together, eats_boarder, eats_group_one, eats_group_two, eats_group_three)
	IF reference_number = "01" THEN
		call navigate_to_screen("STAT", "EATS")
		call create_panel_if_nonexistent
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

Function write_panel_to_MAXIS_EMMA(EMMA_medical_emergency, EMMA_health_consequence, EMMA_verification, EMMA_begin_date, EMMA_end_date)
	call navigate_to_screen("STAT", "EMMA")
	call ERRR_screen_check
	call create_panel_if_nonexistent
	EMWriteScreen EMMA_medical_emergency, 6, 46
	EMWriteScreen EMMA_health_consequence, 8, 46
	EMWriteScreen EMMA_verification, 10, 46
	call create_MAXIS_friendly_date(EMMA_begin_date, 0, 12, 46)
	IF EMMA_end_date <> "" THEN call create_MAXIS_friendly_date(EMMA_end_date, 0, 14, 46)
End function

FUNCTION write_panel_to_MAXIS_EMPS(EMPS_orientation_date, EMPS_orientation_attended, EMPS_good_cause, EMPS_sanc_begin, EMPS_sanc_end, EMPS_memb_at_home, EMPS_care_family, EMPS_crisis, EMPS_hard_employ, EMPS_under1, EMPS_DWP_date)
	call navigate_to_screen("STAT", "EMPS")
	call create_panel_if_nonexistent
	If EMPS_orientation_date <> "" then call create_MAXIS_friendly_date(EMPS_orientation_date, 0, 5, 39) 'enter orientation date
	EMWritescreen left(EMPS_orientation_attended, 1), 5, 65 
	EMWritescreen EMPS_good_cause, 5, 79
	If EMPS_sanc_begin <> "" then call create_MAXIS_friendly_date(EMPS_sanc_begin, 1, 6, 39) 'Sanction begin date
	If EMPS_sanc_end <> "" then call create_MAXIS_friendly_date(EMPS_sanc_end, 1, 6, 65) 'Sanction end date
	EMWritescreen left(EMPS_memb_at_home, 1), 8, 76
	EMWritescreen left(EMPS_care_family, 1), 9, 76
	EMWritescreen left(EMPS_crisis, 1), 10, 76
	EMWritescreen EMPS_hard_employ, 11, 76
	EMWritescreen left(EMPS_under1, 1), 12, 76 'child under 1 exemption
	EMWritescreen "n", 13, 76 'enters n for child under 12 weeks
	If EMPS_DWP_date <> "" then call create_MAXIS_friendly_date(EMPS_DWP_date, 1, 17, 40) 'DWP plan date
	'This populates the child under 1 popup if needed
	IF ucase(left(EMPS_under1, 1)) = "Y" THEN
		EMReadScreen month_to_use, 2, 20, 55
		EMReadScreen start_year, 2, 20, 58
		Emwritescreen "x", 12, 39
		Transmit
		EMReadScreen check_for_blank, 2, 7, 22 'makes sure the popup isn't already filled out
		month_to_use = cint(month_to_use)
		start_year = cint("20" & start_year)
		popup_row = 7 'setting initial starting point for the popup
		popup_col = 22
		IF check_for_blank <> "  " THEN 'blank popup, fill it out!
			FOR i = 1 to 12
				IF month_to_use > 12 THEN 'handling the year change
					popup_month = month_to_use - 12
					year_to_use = start_year +1
				ELSE 
					popup_month = month_to_use
					year_to_use = start_year
				END IF
				IF len(popup_month) = 1 THEN popup_month = "0" & popup_month 'formatting to two digit month
				Emwritescreen popup_month, popup_row, popup_col
				Emwritescreen year_to_use, popup_row, popup_col + 5
				popup_col = popup_col + 11
				month_to_use = month_to_use + 1
				IF popup_col > 55 THEN 'This moves to the next row if necessary
					popup_col = 22
					popup_row = popup_row + 1
				END IF
			NEXT
			PF3 'closing the popup
		END IF
	END IF
End Function

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

'---The custom function to pull FMED information from the Excel file. This function can handle up to 4 FMED rows per client.
FUNCTION write_panel_to_MAXIS_FMED(FMED_medical_mileage, FMED_1_type, FMED_1_verif, FMED_1_ref_num, FMED_1_category, FMED_1_begin, FMED_1_end, FMED_1_amount, FMED_2_type, FMED_2_verif, FMED_2_ref_num, FMED_2_category, FMED_2_begin, FMED_2_end, FMED_2_amount, FMED_3_type, FMED_3_verif, FMED_3_ref_num, FMED_3_category, FMED_3_begin, FMED_3_end, FMED_3_amount, FMED_4_type, FMED_4_verif, FMED_4_ref_num, FMED_4_category, FMED_4_begin, FMED_4_end, FMED_4_amount)
	CALL navigate_to_MAXIS_screen("STAT", "FMED")
	ERRR_screen_check
	EMReadScreen num_of_FMED, 1, 2, 78
	IF num_of_FMED = "0" THEN 
		EMWriteScreen "NN", 20, 79
		transmit
	ELSE
		PF9
	END IF
	
	'Determining where to start writing...
	FMED_row = 9
	DO
		EMReadScreen FMED_available, 2, FMED_row, 25
		IF FMED_available <> "__" THEN FMED_row = FMED_row + 1
		IF FMED_row = 15 THEN 
			PF20
			FMED_row = 9
		END IF
	LOOP UNTIL FMED_available = "__"
	
	IF FMED_1_type <> "" THEN 
		EMWriteScreen FMED_1_type, FMED_row, 25
			IF FMED_1_type = "12" THEN 
				EMReadScreen current_miles, 4, 17, 34
				current_miles = trim(replace(current_miles, "_", " "))
				IF current_miles = "" THEN current_miles = 0
				total_miles = current_miles + FMED_medical_mileage
				EMWriteScreen "    ", 17, 34
				EMWriteScreen total_miles, 17, 34
				FMED_1_verif = "CL"		'Edit: MEDICAL EXPENCE VERIFICATION FOR THIS TYPE CAN ONLY BE CL
				FMED_1_amount = ""		'An FMED amount is not allowed when using Type 12
			END IF
		EMWriteScreen FMED_1_verif, FMED_row, 32
		EMWriteScreen FMED_1_ref_num, FMED_row, 38
		EMWriteScreen FMED_1_category, FMED_row, 44
			FMED_month = DatePart("M", FMED_1_begin)			'Turning the value in FMED_1_begin and FMED_1_end into values the FMED panel can handle.
			IF len(FMED_month) <> 2 THEN FMED_month = "0" & FMED_month
		EMWriteScreen FMED_month, FMED_row, 50
		EMWriteScreen right(DatePart("YYYY", FMED_1_begin), 2), FMED_row, 53
		IF FMED_1_end <> "" THEN 
			FMED_month = DatePart("M", FMED_1_end)
			IF len(FMED_month) <> 2 THEN FMED_month = "0" & FMED_month
			EMWriteScreen FMED_month, FMED_row, 60
			EMWriteScreen right(DatePart("YYYY", FMED_1_end), 2), FMED_row, 63
		END IF
		EMWriteScreen FMED_1_amount, FMED_row, 70
		
		'Next line or new page
		FMED_row = FMED_row + 1
		IF FMED_row = 15 THEN
			PF20
			FMED_row = 9
		END IF
	END IF
	
	IF FMED_2_type <> "" THEN 
		EMWriteScreen FMED_2_type, FMED_row, 25
			IF FMED_2_type = "12" THEN 
				EMReadScreen current_miles, 4, 17, 34
				current_miles = trim(replace(current_miles, "_", " "))
				IF current_miles = "" THEN current_miles = 0			
				total_miles = current_miles + FMED_medical_mileage
				EMWriteScreen "    ", 17, 34
				EMWriteScreen total_miles, 17, 34
				FMED_2_verif = "CL"		'Edit: MEDICAL EXPENCE VERIFICATION FOR THIS TYPE CAN ONLY BE CL
				FMED_2_amount = ""		'An FMED amount is not allowed when using Type 12
			END IF
		EMWriteScreen FMED_2_verif, FMED_row, 32
		EMWriteScreen FMED_2_ref_num, FMED_row, 38
		EMWriteScreen FMED_2_category, FMED_row, 44
			FMED_month = DatePart("M", FMED_2_begin)			'Turning the value in FMED_2_begin and FMED_2_end into values the FMED panel can handle.
			IF len(FMED_month) <> 2 THEN FMED_month = "0" & FMED_month
		EMWriteScreen FMED_month, FMED_row, 50
		EMWriteScreen right(DatePart("YYYY", FMED_2_begin), 2), FMED_row, 53
		IF FMED_2_end <> "" THEN 
			FMED_month = DatePart("M", FMED_2_end)
			IF len(FMED_month) <> 2 THEN FMED_month = "0" & FMED_month
			EMWriteScreen FMED_month, FMED_row, 60
			EMWriteScreen right(DatePart("YYYY", FMED_2_end), 2), FMED_row, 63
		END IF
		EMWriteScreen FMED_2_amount, FMED_row, 70
		
		'Next line or new page
		FMED_row = FMED_row + 1
		IF FMED_row = 15 THEN
			PF20
			FMED_row = 9
		END IF
	END IF
	
	IF FMED_3_type <> "" THEN 
		EMWriteScreen FMED_3_type, FMED_row, 25
			IF FMED_3_type = "12" THEN 
				EMReadScreen current_miles, 4, 17, 34
				current_miles = trim(replace(current_miles, "_", " "))
				IF current_miles = "" THEN current_miles = 0			
				total_miles = current_miles + FMED_medical_mileage
				EMWriteScreen "    ", 17, 34
				EMWriteScreen total_miles, 17, 34
				FMED_3_verif = "CL"		'Edit: MEDICAL EXPENCE VERIFICATION FOR THIS TYPE CAN ONLY BE CL
				FMED_3_amount = ""		'An FMED amount is not allowed when using Type 12
			END IF
		EMWriteScreen FMED_3_verif, FMED_row, 32
		EMWriteScreen FMED_3_ref_num, FMED_row, 38
		EMWriteScreen FMED_3_category, FMED_row, 44
			FMED_month = DatePart("M", FMED_3_begin)			'Turning the value in FMED_3_begin and FMED_3_end into values the FMED panel can handle.
			IF len(FMED_month) <> 2 THEN FMED_month = "0" & FMED_month
		EMWriteScreen FMED_month, FMED_row, 50
		EMWriteScreen right(DatePart("YYYY", FMED_3_begin), 2), FMED_row, 53
		IF FMED_3_end <> "" THEN 
			FMED_month = DatePart("M", FMED_3_end)
			IF len(FMED_month) <> 2 THEN FMED_month = "0" & FMED_month
			EMWriteScreen FMED_month, FMED_row, 60
			EMWriteScreen right(DatePart("YYYY", FMED_3_end), 2), FMED_row, 63
		END IF
		EMWriteScreen FMED_3_amount, FMED_row, 70
		
		'Next line or new page
		FMED_row = FMED_row + 1
		IF FMED_row = 15 THEN
			PF20
			FMED_row = 9
		END IF
	END IF
	
	IF FMED_4_type <> "" THEN 
		EMWriteScreen FMED_4_type, FMED_row, 25
			IF FMED_4_type = "12" THEN 
				EMReadScreen current_miles, 4, 17, 34
				current_miles = trim(replace(current_miles, "_", " "))
				IF current_miles = "" THEN current_miles = 0			
				total_miles = current_miles + FMED_medical_mileage
				EMWriteScreen "    ", 17, 34
				EMWriteScreen total_miles, 17, 34
				FMED_4_verif = "CL"		'Edit: MEDICAL EXPENCE VERIFICATION FOR THIS TYPE CAN ONLY BE CL
				FMED_4_amount = ""		'An FMED amount is not allowed when using Type 12
			END IF
		EMWriteScreen FMED_4_verif, FMED_row, 32
		EMWriteScreen FMED_4_ref_num, FMED_row, 38
		EMWriteScreen FMED_4_category, FMED_row, 44
			FMED_month = DatePart("M", FMED_4_begin)			'Turning the value in FMED_4_begin and FMED_4_end into values the FMED panel can handle.
			IF len(FMED_month) <> 2 THEN FMED_month = "0" & FMED_month
		EMWriteScreen FMED_month, FMED_row, 50
		EMWriteScreen right(DatePart("YYYY", FMED_4_begin), 2), FMED_row, 53
		IF FMED_4_end <> "" THEN 
			FMED_month = DatePart("M", FMED_4_end)
			IF len(FMED_month) <> 2 THEN FMED_month = "0" & FMED_month
			EMWriteScreen FMED_month, FMED_row, 60
			EMWriteScreen right(DatePart("YYYY", FMED_4_end), 2), FMED_row, 63
		END IF
		EMWriteScreen FMED_4_amount, FMED_row, 70
	END IF
	
	transmit
END FUNCTION

Function write_panel_to_MAXIS_HCRE(hcre_appl_addnd_date_input,hcre_retro_months_input,hcre_recvd_by_service_date_input)
	call navigate_to_screen("STAT","HCRE")
	call create_panel_if_nonexistent
	'Converting the Appl Addendum Date into a usable format
	call MAXIS_dater(hcre_appl_addnd_date_input, hcre_appl_addnd_date_output, "HCRE Addendum Date") 
	'Converting the Received by service date into a usable format
	call MAXIS_dater(hcre_recvd_by_service_date_input, hcre_recvd_by_service_date_output, "received by Service Date") 
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

FUNCTION write_panel_to_MAXIS_HEST(HEST_FS_choice_date, HEST_first_month, HEST_heat_air_retro, HEST_electric_retro, HEST_phone_retro, HEST_heat_air_pro, HEST_electric_pro, HEST_phone_pro)
	call navigate_to_screen("STAT", "HEST")
	call create_panel_if_nonexistent
	Emwritescreen "01", 6, 40
	call create_MAXIS_friendly_date(HEST_FS_choice_date, 0, 7, 40)
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
End function

Function write_panel_to_MAXIS_IMIG(IMIG_imigration_status, IMIG_entry_date, IMIG_status_date, IMIG_status_ver, IMIG_status_LPR_adj_from, IMIG_nationality)
	call navigate_to_screen("STAT", "IMIG")
	call ERRR_screen_check
	call create_panel_if_nonexistent
	call create_MAXIS_friendly_date(date, 0, 5, 45)						'Writes actual date, needs to add 2000 as this is weirdly a 4 digit year
	EMWriteScreen datepart("yyyy", date), 5, 51
	EMWriteScreen IMIG_imigration_status, 6, 45							'Writes imig status
	IF IMIG_entry_date <> "" THEN
		call create_MAXIS_friendly_date(IMIG_entry_date, 0, 7, 45)			'Enters year as a 2 digit number, so have to modify manually
		EMWriteScreen datepart("yyyy", IMIG_entry_date), 7, 51
	END IF
	IF IMIG_status_date <> "" THEN
		call create_MAXIS_friendly_date(IMIG_status_date, 0, 7, 71)			'Enters year as a 2 digit number, so have to modify manually
		EMWriteScreen datepart("yyyy", IMIG_status_date), 7, 77
	END IF
	EMWriteScreen IMIG_status_ver, 8, 45								'Enters status ver
	EMWriteScreen IMIG_status_LPR_adj_from, 9, 45						'Enters status LPR adj from
	EMWriteScreen IMIG_nationality, 10, 45								'Enters nationality
	transmit
	transmit
End function

Function write_panel_to_MAXIS_INSA(insa_pers_coop_ohi, insa_good_cause_status, insa_good_cause_cliam_date, insa_good_cause_evidence, insa_coop_cost_effect, insa_insur_name, insa_prescrip_drug_cover, insa_prescrip_end_date, insa_persons_covered)
	call navigate_to_screen("STAT","INSA")
	call create_panel_if_nonexistent
	
	EMWriteScreen insa_pers_coop_ohi, 4, 62
	EMWriteScreen insa_good_cause_status, 5, 62 
	If insa_good_cause_cliam_date <> "" then CALL create_MAXIS_friendly_date(insa_good_cause_cliam_date, 0, 6, 62)
	EMWriteScreen insa_good_cause_evidence, 7, 62
	EMWriteScreen insa_coop_cost_effect, 8, 62
	EMWriteScreen insa_insur_name, 10, 38
	EMWriteScreen insa_prescrip_drug_cover, 11, 62
	If insa_prescrip_end_date <> "" then CALL create_MAXIS_friendly_date(insa_prescrip_end_date, 0, 12, 62)

	'Adding persons covered
	insa_row = 15
	insa_col = 30
	
	insa_persons_covered = replace(insa_persons_covered, " ", "")
	insa_persons_covered = split(insa_persons_covered, ",")
	
	FOR EACH insa_peep IN insa_persons_covered
		EMWriteScreen insa_peep, insa_row, insa_col
		insa_col = insa_col + 4
		IF insa_col = 70 THEN
			insa_col = 30
			insa_row = 16
		END IF
	NEXT
	
	transmit

End Function

FUNCTION write_panel_to_MAXIS_JOBS(jobs_number, jobs_inc_type, jobs_inc_verif, jobs_employer_name, jobs_inc_start, jobs_wkly_hrs, jobs_hrly_wage, jobs_pay_freq)
	call navigate_to_screen("STAT", "JOBS")
	EMWriteScreen reference_number, 20, 76
	EMWriteScreen jobs_number, 20, 79
	transmit
	
	EMReadScreen does_not_exist, 14, 24, 13
	IF does_not_exist = "DOES NOT EXIST" THEN 
		EMWriteScreen "NN", 20, 79
		transmit
	ELSE
		PF9
	END IF
	
	EMWriteScreen jobs_inc_type, 5, 38
	EMWriteScreen jobs_inc_verif, 6, 38
	EMWriteScreen jobs_employer_name, 7, 42
	call create_MAXIS_friendly_date(jobs_inc_start, 0, 9, 35)
	EMWriteScreen jobs_pay_freq, 18, 35
	
	'===== navigates to the SNAP PIC to update the PIC =====
	EMWriteScreen "X", 19, 38
	transmit
	DO
		EMReadScreen at_snap_pic, 12, 3, 22
	LOOP UNTIL at_snap_pic = "Food Support"
	EMReadScreen jobs_pic_wages_per_pp, 7, 17, 57
	EMReadScreen pic_info_exists, 8, 18, 57
	pic_info_exists = trim(pic_info_exists)
	IF pic_info_exists = "" THEN 
		call create_MAXIS_friendly_date(date, 0, 5, 34)
		EMWriteScreen jobs_pay_freq, 5, 64
		EMWriteScreen jobs_wkly_hrs, 8, 64
		EMWriteScreen jobs_hrly_wage, 9, 66
		transmit
		transmit
		EMReadScreen jobs_pic_hrs_per_pp, 6, 16, 51
		EMReadScreen jobs_pic_wages_per_pp, 7, 17, 57
	END IF
	transmit		'<=====navigates out of the PIC
		
	'=====the following bit is for the retrospective & prospective pay dates=====
	EMReadScreen bene_month, 2, 20, 55
	EMReadScreen bene_year, 2, 20, 58
	benefit_month = bene_month & "/01/" & bene_year
	retro_month = DatePart("M", DateAdd("M", -2, benefit_month))
	IF len(retro_month) <> 2 THEN retro_month = "0" & retro_month
	retro_year = right(DatePart("YYYY", DateAdd("M", -2, benefit_month)), 2)
			
	EMWriteScreen retro_month, 12, 25
	EMWriteScreen retro_year, 12, 31
	EMWriteScreen bene_month, 12, 54
	EMWriteScreen bene_year, 12, 60
	
	IF pic_info_exists = "" THEN 		'---If the PIC is blank, the information needs to be added to the main JOBS panel as well.
		EMWriteScreen "05", 12, 28
		EMWriteScreen jobs_pic_wages_per_pp, 12, 38
		EMWriteScreen "05", 12, 57
		EMWriteScreen jobs_pic_wages_per_pp, 12, 67
		EMWriteScreen Int(jobs_pic_hrs_per_pp), 18, 43
		EMWriteScreen Int(jobs_pic_hrs_per_pp), 18, 72
	END IF
		
	IF jobs_pay_freq = 2 OR jobs_pay_freq = 3 THEN
		EMWriteScreen retro_month, 13, 25
		EMWriteScreen retro_year, 13, 31
		EMWriteScreen bene_month, 13, 54
		EMWriteScreen bene_year, 13, 60
		
		IF pic_info_exists = "" THEN 
			EMWriteScreen "19", 13, 28
			EMWriteScreen jobs_pic_wages_per_pp, 13, 38
			EMWriteScreen "19", 13, 57
			EMWriteScreen jobs_pic_wages_per_pp, 13, 67
			EMWriteScreen Int(2 * jobs_pic_hrs_per_pp), 18, 43
			EMWriteScreen Int(2 * jobs_pic_hrs_per_pp), 18, 72
		END IF
	ELSEIF jobs_pay_freq = 4 THEN
		EMWriteScreen retro_month, 13, 25
		EMWriteScreen retro_year, 13, 31
		EMWriteScreen retro_month, 14, 25
		EMWriteScreen retro_year, 14, 31
		EMWriteScreen retro_month, 15, 25
		EMWriteScreen retro_year, 15, 31
		EMWriteScreen bene_month, 13, 54
		EMWriteScreen bene_year, 13, 60
		EMWriteScreen bene_month, 14, 54 
		EMWriteScreen bene_year, 14, 60 
		EMWriteScreen bene_month, 15, 54
		EMWriteScreen bene_year, 15, 60
		
		IF pic_info_exists = "" THEN 
			EMWriteScreen "12", 13, 28
			EMWriteScreen jobs_pic_wages_per_pp, 13, 38
			EMWriteScreen "19", 14, 28
			EMWriteScreen jobs_pic_wages_per_pp, 14, 38
			EMWriteScreen "26", 15, 28
			EMWriteScreen jobs_pic_wages_per_pp, 15, 38
			EMWriteScreen "12", 13, 57
			EMWriteScreen jobs_pic_wages_per_pp, 13, 67
			EMWriteScreen "19", 14, 57 
			EMWriteScreen jobs_pic_wages_per_pp, 14, 67
			EMWriteScreen "26", 15, 57
			EMWriteScreen jobs_pic_wages_per_pp, 15, 67
			EMWriteScreen Int(4 * jobs_pic_hrs_per_pp), 18, 43
			EMWriteScreen Int(4 * jobs_pic_hrs_per_pp), 18, 72
		END IF
	END IF

	'=====determines if the benefit month is current month + 1 and dumps information into the HC income estimator
	IF (bene_month * 1) = (datepart("M", DATE) + 1) THEN		'<===== "bene_month * 1" is needed to convert bene_month from a string to numeric.
		EMWriteScreen "X", 19, 54
		transmit
		
		DO
			EMReadScreen hc_inc_est, 9, 9, 43
		LOOP UNTIL hc_inc_est = "HC Income"
		
		EMWriteScreen jobs_pic_wages_per_pp, 11, 63
		transmit
		transmit
	END IF
END FUNCTION

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

FUNCTION write_panel_to_MAXIS_MMSA(mmsa_liv_arr, mmsa_cont_elig, mmsa_spous_inc, mmsa_shared_hous)
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

Function write_panel_to_MAXIS_MSUR(msur_begin_date)
	call navigate_to_screen("STAT","MSUR")
	call create_panel_if_nonexistent
	
	'msur_begin_date This is the date MSUR began for this client  
	row = 7
	DO
		EMReadScreen available_space, 2, row, 36
		IF available_space = "__" THEN 
			row = row + 1
		ELSE
			EXIT DO
		END IF
	LOOP UNTIL available_space <> "__"
	
	CALL create_MAXIS_friendly_date(msur_begin_date, 0, row, 36)
	Emwritescreen DatePart("YYYY", msure_begin_date), row, 42
	transmit
End Function

'---This function writes using the variables read off of the specialized excel template to the othr panel in MAXIS
Function write_panel_to_MAXIS_OTHR(othr_type, othr_cash_value, othr_cash_value_ver, othr_owed, othr_owed_ver, othr_date, othr_cash_count, othr_SNAP_count, othr_HC_count, othr_IV_count, othr_joint_owner, othr_share_ratio)
	Call navigate_to_screen("STAT", "OTHR")  'navigates to the stat panel
	call create_panel_if_nonexistent
	Emwritescreen othr_type, 6, 40  'enters other asset type
	IF othr_cash_value = "" THEN othr_cash_value = 0
	Emwritescreen othr_cash_value, 8, 40  'enters cash value of asset
	Emwritescreen othr_cash_value_ver, 8, 57  'enters cash value verification code
	IF othr_owed = "" THEN othr_owed = 0
	Emwritescreen othr_owed, 9, 40  'enters amount owed value
	Emwritescreen othr_owed_ver, 9, 57  'enters amount owed verification code
	call create_MAXIS_friendly_date(othr_date, 0, 10, 39)  'enters the as of date in a MAXIS friendly format. mm/dd/yy
	Emwritescreen othr_cash_count, 12, 50  'enters y/n if counted for cash
	Emwritescreen othr_SNAP_count, 12, 57  'enters y/n if counted for snap
	Emwritescreen othr_HC_count, 12, 64  'enters y/n if counted for hc
	Emwritescreen othr_IV_count, 12, 73  'enters y/n if counted for iv
	Emwritescreen othr_joint_owner, 13, 44  'enters if it is a jointly owned other asset
	Emwritescreen left(othr_share_ratio, 1), 15, 50  'enters the ratio of ownership using the left 1 digit of what is entered into the file
	Emwritescreen right(othr_share_ratio, 1), 15, 54  'enters the ratio of ownership using the right 1 digit of what is entered into the file
End Function

FUNCTION write_panel_to_MAXIS_PARE(appl_date, PARE_child_1, PARE_child_1_relation, PARE_child_1_verif, PARE_child_2, PARE_child_2_relation, PARE_child_2_verif, PARE_child_3, PARE_child_3_relation, PARE_child_3_verif, PARE_child_4, PARE_child_4_relation, PARE_child_4_verif, PARE_child_5, PARE_child_5_relation, PARE_child_5_verif, PARE_child_6, PARE_child_6_relation, PARE_child_6_verif)
	Call navigate_to_screen("STAT", "PARE") 
	call create_panel_if_nonexistent
	CALL create_MAXIS_friendly_date(appl_date, 0, 5, 37)
	EMWriteScreen DatePart("YYYY", appl_date), 5, 43
	
	IF len(PARE_child_1) = 1 THEN PARE_child_1 = "0" & PARE_child_1
	IF len(PARE_child_2) = 1 THEN PARE_child_1 = "0" & PARE_child_2
	IF len(PARE_child_3) = 1 THEN PARE_child_1 = "0" & PARE_child_3
	IF len(PARE_child_4) = 1 THEN PARE_child_1 = "0" & PARE_child_4
	IF len(PARE_child_5) = 1 THEN PARE_child_1 = "0" & PARE_child_5
	IF len(PARE_child_6) = 1 THEN PARE_child_1 = "0" & PARE_child_6
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

'---This function writes using the variables read off of the specialized excel template to the pben panel in MAXIS
Function write_panel_to_MAXIS_PBEN(pben_referal_date, pben_type, pben_appl_date, pben_appl_ver, pben_IAA_date, pben_disp)
	Call navigate_to_screen("STAT", "PBEN")  'navigates to the stat panel
	call create_panel_if_nonexistent
	Emreadscreen pben_row_check, 2, 8, 24  'reads the MAXIS screen to find out if the PBEN row has already been used. 
	If pben_row_check = "  " THEN   'if the row is blank it enters it in the 8th row.
		Emwritescreen pben_type, 8, 24  'enters pben type code
		call create_MAXIS_friendly_date(pben_referal_date, 0, 8, 40)  'enters referal date in MAXIS friendly format mm/dd/yy
		call create_MAXIS_friendly_date(pben_appl_date, 0, 8, 51)  'enters appl date in  MAXIS friendly format mm/dd/yy
		Emwritescreen pben_appl_ver, 8, 62  'enters appl verification code
		call create_MAXIS_friendly_date(pben_IAA_date, 0, 8, 66)  'enters IAA date in MAXIS friendly format mm/dd/yy
		Emwritescreen pben_disp, 8, 77  'enters the status of pben application 
	else 
		EMreadscreen pben_row_check, 2, 9, 24  'if row 8 is filled already it will move to row 9 and see if it has been used. 
		IF pben_row_check = "  " THEN  'if the 9th row is blank it enters the information there. 
		'second pben row
			Emwritescreen pben_type, 9, 24
			call create_MAXIS_friendly_date(pben_referal_date, 0, 9, 40)
			call create_MAXIS_friendly_date(pben_appl_date, 0, 9, 51)
			Emwritescreen pben_appl_ver, 9, 62
			call create_MAXIS_friendly_date(pben_IAA_date, 0, 9, 66)
			Emwritescreen pben_disp, 9, 77
		else
		Emreadscreen pben_row_check, 2, 10, 24  'if row 8 is filled already it will move to row 9 and see if it has been used.
			IF pben-row_check = "  " THEN  'if the 9th row is blank it enters the information there.
			'third pben row
				Emwritescreen pben_type, 10, 24
				call create_MAXIS_friendly_date(pben_referal_date, 0, 10, 40)
				call create_MAXIS_friendly_date(pben_appl_date, 0, 10, 51)
				Emwritescreen pben_appl_ver, 10, 62
				call create_MAXIS_friendly_date(pben_IAA_date, 0, 10, 66)
				Emwritescreen pben_disp, 10, 77
			END IF
		END IF
	END IF
End Function

Function write_panel_to_MAXIS_PDED(PDED_wid_deduction, PDED_adult_child_disregard, PDED_wid_disregard, PDED_unea_income_deduction_reason, PDED_unea_income_deduction_value, PDED_earned_income_deduction_reason, PDED_earned_income_deduction_value, PDED_ma_epd_inc_asset_limit, PDED_guard_fee, PDED_rep_payee_fee, PDED_other_expense, PDED_shel_spcl_needs, PDED_excess_need, PDED_restaurant_meals)
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

FUNCTION write_panel_to_MAXIS_PREG(PREG_conception_date, PREG_conception_date_ver, PREG_third_trimester_ver, PREG_due_date, PREG_multiple_birth)
	call navigate_to_screen("STAT", "PREG")
	call create_panel_if_nonexistent
	EMWritescreen "NN", 20, 79
	transmit
	call create_MAXIS_friendly_date(PREG_conception_date, 0, 6, 53)
	third_trimester_date = dateadd("M", 6, PREG_conception_date)
	CALL create_MAXIS_friendly_date(third_trimester_date, 0, 8, 53)
	call create_MAXIS_friendly_date(PREG_due_date, 1, 10, 53)
	EMWritescreen PREG_conception_date_ver, 6, 75
	EMWritescreen PREG_third_trimester_ver, 8, 75
	EMWritescreen PREG_multiple_birth, 14, 53
	transmit
end function

'---This function writes using the variables read off of the specialized excel template to the rbic panel in MAXIS
Function write_panel_to_MAXIS_RBIC(rbic_type, rbic_start_date, rbic_end_date, rbic_group_1, rbic_retro_income_group_1, rbic_prosp_income_group_1, rbic_ver_income_group_1, rbic_group_2, rbic_retro_income_group_2, rbic_prosp_income_group_2, rbic_ver_income_group_2, rbic_group_3, rbic_retro_income_group_3, rbic_prosp_income_group_3, rbic_ver_income_group_3, rbic_retro_hours, rbic_prosp_hours, rbic_exp_type_1, rbic_exp_retro_1, rbic_exp_prosp_1, rbic_exp_ver_1, rbic_exp_type_2, rbic_exp_retro_2, rbic_exp_prosp_2, rbic_exp_ver_2)
	call navigate_to_screen("STAT", "RBIC")  'navigates to the stat panel
	call create_panel_if_nonexistent
	EMwritescreen rbic_type, 5, 44  'enters rbic type code
	call create_MAXIS_friendly_date(rbic_start_date, 0, 6, 44)  'creates and enters a MAXIS friend date in the format mm/dd/yy for rbic start date
	IF rbic_end_date <> "" THEN call create_MAXIS_friendly_date(rbic_end_date, 6, 68)  'creates and enters a MAXIS friend date in the format mm/dd/yy for rbic end date
	rbic_group_1 = replace(rbic_group_1, " ", "")  'this will replace any spaces in the array with nothing removing the spaces.
	rbic_group_1 = split(rbic_group_1, ",")  'this will split up the reference numbers in the array based on commas
	rbic_col = 25                            'this will set the starting column to enter rbic reference numbers
	For each rbic_hh_memb in rbic_group_1    'for each reference number that is in the array for group 1 it will enter the reference numbers into the correct row.
		EMwritescreen rbic_hh_memb, 10, rbic_col
		rbic_col = rbic + 3
	NEXT
	EMwritescreen rbic_retro_income_group_1, 10, 47  'enters the rbic retro income for group 1
	EMwritescreen rbic_prosp_income_group_1, 10, 62  'enters the rbic prospective income for group 1
	EMwritescreen rbic_ver_income_group_1, 10, 76    'enters the income verification code for group 1
	rbic_group_2 = replace(rbic_group_2, " ", "")
	rbic_group_2 = split(rbic_group_2, ",")
	rbic_col = 25
	For each rbic_hh_memb in rbic_group_2    'for each reference number that is in the array for group 2 it will enter the reference numbers into the correct row.
		EMwritescreen rbic_hh_memb, 11, rbic_col
		rbic_col = rbic + 3
	NEXT
	EMwritescreen rbic_retro_income_group_2, 11, 47  'enters the rbic retro income for group 2
	EMwritescreen rbic_prosp_income_group_2, 11, 62  'enters the rbic prospective income for group 2
	EMwritescreen rbic_ver_income_group_2, 11, 76    'enters the income verification code for group 2
	rbic_group_3 = replace(rbic_group_3, " ", "")
	rbic_group_3 = split(rbic_group_3, ",")
	rbic_col = 25
	For each rbic_hh_memb in rbic_group_3    'for each reference number that is in the array for group 3 it will enter the reference numbers into the correct row.
		EMwritescreen rbic_hh_memb, 10, rbic_col
		rbic_col = rbic + 3
	NEXT
	EMwritescreen rbic_retro_income_group_3, 12, 47  'enters the rbic retro income for group 3
	EMwritescreen rbic_prosp_income_group_3, 12, 62  'enters the rbic prospective income for group 3
	EMwritescreen rbic_ver_income_group_3, 12, 76    'enters the income verification code for group 3
	EMwritescreen rbic_retro_hours, 13, 52  'enters the retro hours
	EMwritescreen rbic_prosp_hours, 13, 67  'enters the prospective hours
	EMwritescreen rbic_exp_type_1, 15, 25   'enters the expenses type for group 1
	EMwritescreen rbic_exp_retro_1, 15, 47  'enters the expenses retro for group 1
	EMwritescreen rbic_exp_prosp_1, 15, 62  'enters the expenses prospective for group 1
	EMwritescreen rbic_exp_ver_1, 15, 76    'enters the expenses verification code for group 1
	EMwritescreen rbic_exp_type_2, 16, 25   'enters the expenses type for group 2
	EMwritescreen rbic_exp_retro_2, 16, 47  'enters the expenses retro for group 2
	EMwritescreen rbic_exp_prosp_2, 16, 62  'enters the expenses prospective for group 2
	EMwritescreen rbic_exp_ver_2, 16, 76    'enters the expenses verification code for group 2
end function

'---This function writes using the variables read off of the specialized excel template to the rest panel in MAXIS
Function write_panel_to_MAXIS_REST(rest_type, rest_type_ver, rest_market, rest_market_ver, rest_owed, rest_owed_ver, rest_date, rest_status, rest_joint, rest_share_ratio, rest_agreement_date)
	Call navigate_to_screen("STAT", "REST")  'navigates to the stat panel
	call create_panel_if_nonexistent
	Emwritescreen rest_type, 6, 39  'enters residence type
	Emwritescreen rest_type_ver, 6, 62  'enters verification of residence type
	Emwritescreen rest_market, 8, 41  'enters market value of residence
	Emwritescreen rest_market_ver, 8, 62  'enters market value verification code
	Emwritescreen rest_owed, 9, 41  'enters amount owned on residence
	Emwritescreen rest_owed_ver, 9, 62  'enters amount owed verification code
	call create_MAXIS_friendly_date(rest_date, 0, 10, 39)  'enters the as of date in a MAXIS friendly format. mm/dd/yy
	Emwritescreen rest_status, 12, 54  'enters property status code
	Emwritescreen rest_joint, 13, 54  'enters if it is a jointly owned home
	Emwritescreen left(rest_share_ratio, 1), 14, 54  'enters the ratio of ownership using the left 1 digit of what is entered into the file
	Emwritescreen right(rest_share_ratio, 1), 14, 58  'enters the ratio of ownership using the right 1 digit of what is entered into the file
	IF rest_agreement_date <> "" THEN call create_MAXIS_friendly_date(rest_agreement_date, 0, 16, 62)
End Function

Function write_panel_to_MAXIS_SCHL(appl_date, SCHL_status, SCHL_ver, SCHL_type, SCHL_district_nbr, SCHL_kindergarten_start_date, SCHL_grad_date, SCHL_grad_date_ver, SCHL_primary_secondary_funding, SCHL_FS_eligibility_status, SCHL_higher_ed)
	EMWriteScreen "SCHL", 20, 71
	EMWriteScreen reference_number, 20, 76
	transmit
	
	EMReadScreen num_of_SCHL, 1, 2, 78
	IF num_of_SCHL = "0" THEN
		EMWriteScreen "NN", 20, 79
		transmit
	
		call create_MAXIS_friendly_date(appl_date, 0, 5, 40)						'Writes actual date, needs to add 2000 as this is weirdly a 4 digit year
		EMWriteScreen datepart("yyyy", appl_date), 5, 46
		EMWriteScreen SCHL_status, 6, 40
		EMWriteScreen SCHL_ver, 6, 63
		EMWriteScreen SCHL_type, 7, 40
		IF len(SCHL_district_nbr) <> 4 THEN
			DO
				SCHL_district_nbr = "0" & SCHL_district_nbr
			LOOP UNTIL len(SCHL_district_nbr) = 4
		END IF
		EMWriteScreen SCHL_district_nbr, 8, 40
		If SCHL_kindergarten_start_date <> "" then call create_MAXIS_friendly_date(SCHL_kindergarten_start_date, 0, 10, 63)
		EMWriteScreen left(SCHL_grad_date, 2), 11, 63
		EMWriteScreen right(SCHL_grad_date, 2), 11, 66
		EMWriteScreen SCHL_grad_date_ver, 12, 63
		EMWriteScreen SCHL_primary_secondary_funding, 14, 63
		EMWriteScreen SCHL_FS_eligibility_status, 16, 63
		EMWriteScreen SCHL_higher_ed, 18, 63
		transmit
	END IF
End function

'---This function writes using the variables read off of the specialized excel template to the secu panel in MAXIS
Function write_panel_to_MAXIS_SECU(secu_type, secu_pol_numb, secu_name, secu_cash_val, secu_date, secu_cash_ver, secu_face_val, secu_withdraw, secu_cash_count, secu_SNAP_count, secu_HC_count, secu_GRH_count, secu_IV_count, secu_joint, secu_share_ratio)
	Call navigate_to_screen("STAT", "SECU")  'navigates to the stat panel
	call create_panel_if_nonexistent
	Emwritescreen secu_type, 6, 50  'enters security type
	Emwritescreen secu_pol_numb, 7, 50  'enters policy number
	Emwritescreen secu_name, 8, 50  'enters name of policy
	Emwritescreen secu_cash_val, 10, 52  'enters cash value of policy
	call create_MAXIS_friendly_date(secu_date, 0, 11, 35)  'enters the as of date in a MAXIS friendly format. mm/dd/yy
	Emwritescreen secu_cash_ver, 11, 50  'enters cash value verification code
	Emwritescreen secu_face_val, 12, 52  'enters face value of policy
	Emwritescreen secu_withdraw, 13, 52  'enters withdrawl penalty
	Emwritescreen secu_cash_count, 15, 50  'enters y/n if counted for cash
	Emwritescreen secu_SNAP_count, 15, 57  'enters y/n if counted for snap
	Emwritescreen secu_HC_count, 15, 64  'enters y/n if counted for hc
	Emwritescreen secu_GRH_count, 15, 72  'enters y/n if counted for grh
	Emwritescreen secu_IV_count, 15, 80  'enters y/n if counted for iv
	Emwritescreen secu_joint, 16, 44  'enters if it is a jointly owned security
	Emwritescreen left(secu_share_ratio, 1), 16, 76  'enters the ratio of ownership using the left 1 digit of what is entered into the file
	Emwritescreen right(secu_share_ratio, 1), 16, 80  'enters the ratio of ownership using the right 1 digit of what is entered into the file
End Function

FUNCTION write_panel_to_MAXIS_SHEL(SHEL_subsidized, SHEL_shared, SHEL_paid_to, SHEL_rent_retro, SHEL_rent_retro_ver, SHEL_rent_pro, SHEL_rent_pro_ver, SHEL_lot_rent_retro, SHEL_lot_rent_retro_ver, SHEL_lot_rent_pro, SHEL_lot_rent_pro_ver, SHEL_mortgage_retro, SHEL_mortgage_retro_ver, SHEL_mortgage_pro, SHEL_mortgage_pro_ver, SHEL_insur_retro, SHEL_insur_retro_ver, SHEL_insur_pro, SHEL_insur_pro_ver, SHEL_taxes_retro, SHEL_taxes_retro_ver, SHEL_taxes_pro, SHEL_taxes_pro_ver, SHEL_room_retro, SHEL_room_retro_ver, SHEL_room_pro, SHEL_room_pro_ver, SHEL_garage_retro, SHEL_garage_retro_ver, SHEL_garage_pro, SHEL_garage_pro_ver, SHEL_subsidy_retro, SHEL_subsidy_retro_ver, SHEL_subsidy_pro, SHEL_subsidy_pro_ver)
	call navigate_to_screen("STAT", "SHEL")
	call create_panel_if_nonexistent
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
end function

FUNCTION write_panel_to_MAXIS_SIBL(SIBL_group_1, SIBL_group_2, SIBL_group_3)
	call navigate_to_screen("STAT", "SIBL")
	EMReadScreen num_of_SIBL, 1, 2, 78
	IF num_of_SIBL = "0" THEN 
		EMWriteScreen "NN", 20, 79
		transmit
	END IF
		
	If SIBL_group_1 <> "" then 
		EMWritescreen "01", 7, 28
		SIBL_group_1 = replace(SIBL_group_1, " ", "") 'Removing spaces
		SIBL_group_1 = split(SIBL_group_1, ",") 'Splits the sibling group value into an array by commas
		SIBL_col = 39
		For Each SIBL_group_member in SIBL_group_1 'Writes the member numbers onto the group line
			EMWritescreen SIBL_group_member, 7, SIBL_col
			SIBL_col = SIBL_col + 4
		Next
	End if
	
	If SIBL_group_2 <> "" then
		EMWritescreen "02", 8, 28
		SIBL_group_2 = replace(SIBL_group_2, " ", "")
		SIBL_group_2 = split(SIBL_group_2, ",")
		SIBL_col = 39
		For Each SIBL_group_member in SIBL_group_2
			EMWritescreen SIBL_group_member, 8, SIBL_col
			SIBL_col = SIBL_col + 4
		Next
	End if
	
	If SIBL_group_3 <> "" then
		EMWritescreen "03", 9, 28
		SIBL_group_2 = replace(SIBL_group_3, " ", "")
		SIBL_group_2 = split(SIBL_group_3, ",")
		SIBL_col = 39
		For Each SIBL_group_member in SIBL_group_3
			EMWritescreen SIBL_group_member, 9, SIBL_col
			SIBL_col = SIBL_col + 4
		Next
	End if		
	transmit
end function

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

Function write_panel_to_MAXIS_STEC(STEC_type_1, STEC_amt_1, STEC_actual_from_thru_months_1, STEC_ver_1, STEC_earmarked_amt_1, STEC_earmarked_from_thru_months_1, STEC_type_2, STEC_amt_2, STEC_actual_from_thru_months_2, STEC_ver_2, STEC_earmarked_amt_2, STEC_earmarked_from_thru_months_2)
	EMWriteScreen "STEC", 20, 71
	EMWriteSCreen reference_number, 20, 76
	transmit
	
	EMReadScreen num_of_STEC, 1, 2, 78
	IF num_of_STEC = "0" THEN
		EMWriteScreen "NN", 20, 79
		transmit
	
		EMWriteScreen STEC_type_1, 8, 25				'STEC 1
		EMWriteScreen STEC_amt_1, 8, 31
		STEC_actual_from_thru_months_1 = replace(STEC_actual_from_thru_months_1, " ", "")
		EMWriteScreen left(STEC_actual_from_thru_months_1, 2), 8, 41
		EMWriteScreen mid(STEC_actual_from_thru_months_1, 4, 2), 8, 44
		EMWriteScreen mid(STEC_actual_from_thru_months_1, 7, 2), 8, 48
		EMWriteScreen right(STEC_actual_from_thru_months_1, 2), 8, 51
		EMWriteScreen STEC_ver_1, 8, 55
		EMWriteScreen STEC_earmarked_amt_1, 8, 59
		STEC_earmarked_from_thru_months_1 = replace(STEC_earmarked_from_thru_months_1, " ", "")
		EMWriteScreen left(STEC_earmarked_from_thru_months_1, 2), 8, 69
		EMWriteScreen mid(STEC_earmarked_from_thru_months_1, 4, 2), 8, 72
		EMWriteScreen mid(STEC_earmarked_from_thru_months_1, 7, 2), 8, 76
		EMWriteScreen right(STEC_earmarked_from_thru_months_1, 2), 8, 79
		EMWriteScreen STEC_type_2, 9, 25				'STEC 1
		EMWriteScreen STEC_amt_2, 9, 31
		STEC_actual_from_thru_months_2 = replace(STEC_actual_from_thru_months_2, " ", "")
		EMWriteScreen left(STEC_actual_from_thru_months_2, 2), 9, 41
		EMWriteScreen mid(STEC_actual_from_thru_months_2, 4, 2), 9, 44
		EMWriteScreen mid(STEC_actual_from_thru_months_2, 7, 2), 9, 48
		EMWriteScreen right(STEC_actual_from_thru_months_2, 2), 9, 51
		EMWriteScreen STEC_ver_2, 9, 55
		EMWriteScreen STEC_earmarked_amt_2, 9, 59
		STEC_earmarked_from_thru_months_2 = replace(STEC_earmarked_from_thru_months_2, " ", "")
		EMWriteScreen left(STEC_earmarked_from_thru_months_2, 2), 9, 69
		EMWriteScreen mid(STEC_earmarked_from_thru_months_2, 4, 2), 9, 72
		EMWriteScreen mid(STEC_earmarked_from_thru_months_2, 7, 2), 9, 76
		EMWriteScreen right(STEC_earmarked_from_thru_months_2, 2), 9, 79
		transmit
	END IF
End function

Function write_panel_to_MAXIS_STIN(STIN_type_1, STIN_amt_1, STIN_avail_date_1, STIN_months_covered_1, STIN_ver_1, STIN_type_2, STIN_amt_2, STIN_avail_date_2, STIN_months_covered_2, STIN_ver_2)
	EMWriteScreen "STIN", 20, 71
	EMWriteSCreen reference_number, 20, 76
	transmit
	
	EMReadScreen num_of_STIN, 1, 2, 78
	IF num_of_STIN = "0" THEN
		EMWriteScreen "NN", 20, 79
		transmit
		
		EMWriteScreen STIN_type_1, 8, 27				'STIN 1
		EMWriteScreen STIN_amt_1, 8, 34
		call create_MAXIS_friendly_date(STIN_avail_date_1, 0, 8, 46)
		STIN_months_covered_1 = replace(STIN_months_covered_1, " ", "")
		EMWriteScreen left(STIN_months_covered_1, 2), 8, 58
		EMWriteScreen mid(STIN_months_covered_1, 4, 2), 8, 61
		EMWriteScreen mid(STIN_months_covered_1, 7, 2), 8, 67
		EMWriteScreen right(STIN_months_covered_1, 2), 8, 70
		EMWriteScreen STIN_ver_1, 8, 76
		EMWriteScreen STIN_type_2, 9, 27				'STIN 2
		EMWriteScreen STIN_amt_2, 9, 34
		STIN_avail_date_2 = replace(STIN_avail_date_2, " ", "")
		IF STIN_avail_date_2 <> "" THEN call create_MAXIS_friendly_date(STIN_avail_date_2, 0, 9, 46)
		EMWriteScreen left(STIN_months_covered_2, 2), 9, 58
		EMWriteScreen mid(STIN_months_covered_2, 4, 2), 9, 61
		EMWriteScreen mid(STIN_months_covered_2, 7, 2), 9, 67
		EMWriteScreen right(STIN_months_covered_2, 2), 9, 70
		EMWriteScreen STIN_ver_2, 9, 76
		transmit
	END IF
End function

Function write_panel_to_MAXIS_STWK(STWK_empl_name, STWK_wrk_stop_date, STWK_wrk_stop_date_verif, STWK_inc_stop_date, STWK_refused_empl_yn, STWK_vol_quit, STWK_ref_empl_date, STWK_gc_cash, STWK_gc_grh, STWK_gc_fs, STWK_fs_pwe, STWK_maepd_ext)
	call navigate_to_screen("STAT","STWK")
	call create_panel_if_nonexistent
	
	EMWriteScreen stwk_empl_name, 6, 46
	If stwk_wrk_stop_date <> "" then CALL create_MAXIS_friendly_date(stwk_wrk_stop_date, 0, 7, 46)
	EMWriteScreen stwk_wrk_stop_date_verif, 7, 63
	IF stwk_inc_stop_date <> "" THEN CALL create_MAXIS_friendly_date(stwk_inc_stop_date, 0, 8, 46)
	EMWriteScreen stwk_refused_empl_yn, 8, 78
	EMWriteScreen stwk_vol_quit, 10, 46
	If stwk_ref_empl_date <> "" then CALL create_MAXIS_friendly_date(stwk_ref_empl_date, 0, 10, 72)
	EMWriteScreen stwk_gc_cash, 12, 52
	EMWriteScreen stwk_gc_grh, 12, 60
	EMWriteScreen stwk_gc_fs, 12, 67
	EMWriteScreen stwk_fs_pwe, 14, 46
	EMWriteScreen stwk_maepd_ext, 16, 46
	Transmit
End Function

FUNCTION write_panel_to_MAXIS_TYPE_PROG_REVW(appl_date, type_cash_yn, type_hc_yn, type_fs_yn, prog_mig_worker, revw_ar_or_ir, revw_exempt)
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
				call create_MAXIS_friendly_date(appl_date, 0, 6, 33)
				call create_MAXIS_friendly_date(appl_date, 0, 6, 44)
				call create_MAXIS_friendly_date(appl_date, 0, 6, 55)
			END IF
			IF type_fs_yn = "Y" THEN
				call create_MAXIS_friendly_date(appl_date, 0, 10, 33)
				call create_MAXIS_friendly_date(appl_date, 0, 10, 44)
				call create_MAXIS_friendly_date(appl_date, 0, 10, 55)
			END IF
			IF type_hc_yn = "Y" THEN
				call create_MAXIS_friendly_date(appl_date, 0, 12, 33)
				call create_MAXIS_friendly_date(appl_date, 0, 12, 55)
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
				call create_MAXIS_friendly_date(cash_review_date, 0, 9, 37)
			END IF
			IF type_fs_yn = "Y" THEN
				EMWriteScreen "X", 5, 58
				transmit
				DO
					EMReadScreen food_support_reports, 20, 5, 30
				LOOP UNTIL food_support_reports = "FOOD SUPPORT REPORTS"
				fs_csr_date = dateadd("M", 6, appl_date)
				fs_er_date = dateadd("M", 12, appl_date)
				call create_MAXIS_friendly_date(fs_csr_date, 0, 9, 26)
				call create_MAXIS_friendly_date(fs_er_date, 0, 9, 64)
				transmit
			END IF
			IF type_hc_yn = "Y" THEN
				EMWriteScreen "X", 5, 71
				transmit
				DO
					EMReadScreen health_care_renewals, 20, 4, 32
				LOOP UNTIL health_care_renewals = "HEALTH CARE RENEWALS"
				IF revw_ar_or_ir = "AR" THEN
					call create_MAXIS_friendly_date((dateadd("M", 6, appl_date)), 0, 8, 71)
				ELSEIF revw_ar_or_ir = "IR" THEN
					call create_MAXIS_friendly_date((dateadd("M", 6, appl_date)), 0, 8, 27)
				END IF
				call create_MAXIS_friendly_date((dateadd("M", 12, appl_date)), 0, 9, 27)
				EMWriteScreen revw_exempt, 9, 71
				transmit
			END IF
	END IF
END FUNCTION

FUNCTION write_panel_to_MAXIS_UNEA(unea_number, unea_inc_type, unea_inc_verif, unea_claim_suffix, unea_start_date, unea_pay_freq, unea_inc_amount, ssn_first, ssn_mid, ssn_last)
	call navigate_to_screen("STAT", "UNEA")
	PF10
	EMWriteScreen reference_number, 20, 76
	EMWriteScreen unea_number, 20, 79
	transmit
	
	EMReadScreen does_not_exist, 14, 24, 13
	IF does_not_exist = "DOES NOT EXIST" THEN
		EMWriteScreen "NN", 20, 79
		transmit
		
		'Putting this part in with the NN because otherwise the script will update it in later months and change claim number information.
		EMWriteScreen unea_inc_type, 5, 37
		EMWriteScreen unea_inc_verif, 5, 65
		EMWriteScreen (ssn_first & ssn_mid & ssn_last & unea_claim_suffix), 6, 37
		call create_MAXIS_friendly_date(unea_start_date, 0, 7, 37)
	ELSE
		PF9
	END IF

	'=====Navigates to the PIC for UNEA=====
	EMWriteScreen "X", 10, 26
	transmit
	EMReadScreen pic_info_exists, 6, 18, 58		'---Deteremining if PIC info exists. If it does, the script will just back out.
	pic_info_exists = trim(pic_info_exists)
	IF pic_info_exists = "" THEN
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
	ELSE
		PF3
	END IF

	'=====the following bit is for the retrospective & prospective pay dates=====
	EMReadScreen bene_month, 2, 20, 55
	EMReadScreen bene_year, 2, 20, 58
	current_bene_month = bene_month & "/01/" & bene_year
	retro_month = datepart("M", DateAdd("M", -2, current_bene_month))
	IF len(retro_month) <> 2 THEN retro_month = "0" & retro_month
	retro_year = right(datepart("YYYY", DateAdd("M", -2, current_bene_month)), 2)
	
	EMWriteScreen retro_month, 13, 25
	EMWriteScreen retro_year, 13, 31
	EMWriteScreen bene_month, 13, 54
	EMWriteScreen bene_year, 13, 60
	
	IF pic_info_exists = "" THEN 	'---Meaning, the case has PIC info...which is to say that this is a PF9 and not a NN
		EMWriteScreen "05", 13, 28
		EMWriteScreen unea_inc_amount, 13, 39
		EMWriteScreen "05", 13, 57
		EMWriteScreen unea_inc_amount, 13, 68	
	END IF
	
	IF unea_pay_freq = "2" OR unea_pay_freq = "3" THEN
		EMWriteScreen retro_month, 14, 25
		EMWriteScreen retro_year, 14, 31
		EMWriteScreen bene_month, 14, 54
		EMWriteScreen bene_year, 14, 60
				
		IF pic_info_exists = "" THEN 
			EMWriteScreen "19", 14, 28
			EMWriteScreen "19", 14, 57
			EMWriteScreen unea_inc_amount, 14, 39
			EMWriteScreen unea_inc_amount, 14, 68
		END IF
	ELSEIF unea_pay_freq = "4" THEN
		EMWriteScreen retro_month, 14, 25
		EMWriteScreen retro_year, 14, 31
		EMWriteScreen retro_month, 15, 25
		EMWriteScreen retro_year, 15, 31
		EMWriteScreen retro_month, 16, 25
		EMWriteScreen retro_year, 16, 31
		EMWriteScreen bene_month, 14, 54
		EMWriteScreen bene_year, 14, 60
		EMWriteScreen bene_month, 15, 54 
		EMWriteScreen bene_year, 15, 60 
		EMWriteScreen bene_month, 16, 54
		EMWriteScreen bene_year, 16, 60
		
		IF pic_info_exists = "" THEN 
			EMWriteScreen "12", 14, 28
			EMWriteScreen unea_inc_amount, 14, 39
			EMWriteScreen "19", 15, 28
			EMWriteScreen unea_inc_amount, 15, 39
			EMWriteScreen "26", 16, 28
			EMWriteScreen unea_inc_amount, 16, 39
			EMWriteScreen "12", 14, 57
			EMWriteScreen unea_inc_amount, 14, 68
			EMWriteScreen "19", 15, 57 
			EMWriteScreen unea_inc_amount, 15, 68 
			EMWriteScreen "26", 16, 57
			EMWriteScreen unea_inc_amount, 16, 68
		END IF
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

FUNCTION write_panel_to_MAXIS_WKEX(program, fed_tax_retro, fed_tax_prosp, fed_tax_verif, state_tax_retro, state_tax_prosp, state_tax_verif, fica_retro, fica_prosp, fica_verif, tran_retro, tran_prosp, tran_verif, tran_imp_rel, meals_retro, meals_prosp, meals_verif, meals_imp_rel, uniforms_retro, uniforms_prosp, uniforms_verif, uniforms_imp_rel, tools_retro, tools_prosp, tools_verif, tools_imp_rel, dues_retro, dues_prosp, dues_verif, dues_imp_rel, othr_retro, othr_prosp, othr_verif, othr_imp_rel, HC_Exp_Fed_Tax, HC_Exp_State_Tax, HC_Exp_FICA, HC_Exp_Tran, HC_Exp_Tran_imp_rel, HC_Exp_Meals, HC_Exp_Meals_Imp_Rel, HC_Exp_Uniforms, HC_Exp_Uniforms_Imp_Rel, HC_Exp_Tools, HC_Exp_Tools_Imp_Rel, HC_Exp_Dues, HC_Exp_Dues_Imp_Rel, HC_Exp_Othr, HC_Exp_Othr_Imp_Rel)
	CALL navigate_to_MAXIS_screen("STAT", "WKEX")
	ERRR_screen_check
	
	EMWriteScreen reference_number, 20, 76
	transmit
	
	'Determining the number of WKEX panels so the script knows how to handle the incoming information.
	EMReadScreen num_of_WKEX_panels, 1, 2, 78
	IF num_of_WKEX_panels = "5" THEN		'If there are already 5 WKEX panels, the script will not create a new panel.
		EXIT FUNCTION 
	ELSEIF num_of_WKEX_panels = "0" THEN		
		EMWriteScreen "__", 20, 76
		EMWriteScreen "NN", 20, 79
		transmit
		
		'---When the script needs to generate a new WKEX, it will enter the information for that panel...
		EMWriteScreen program, 5, 33
		EMWriteScreen fed_tax_retro, 7, 43
		EMWriteScreen fed_tax_prosp, 7, 57
		EMWriteScreen fed_tax_verif, 7, 69
		EMWriteScreen state_tax_retro, 8, 43
		EMWriteScreen state_tax_prosp, 8, 57
		EMWriteScreen state_tax_verif, 8, 69
		EMWriteScreen fica_retro, 9, 43
		EMWriteScreen fica_prosp, 9, 57
		EMWriteScreen fica_verif, 9, 69
		EMWriteScreen tran_retro, 10, 43
		EMWriteScreen tran_prosp, 10, 57
		EMWriteScreen tran_verif, 10, 69
		EMWriteScreen tran_imp_rel, 10, 75
		EMWriteScreen meals_retro, 11, 43
		EMWriteScreen meals_prosp, 11, 57
		EMWriteScreen meals_verif, 11, 69
		EMWriteScreen meals_imp_rel, 11, 75
		EMWriteScreen uniforms_retro, 12, 43
		EMWriteScreen uniforms_prosp, 12, 57
		EMWriteScreen uniforms_verif, 12, 69
		EMWriteScreen uniforms_imp_rel, 12, 75
		EMWriteScreen tools_retro, 13, 43
		EMWriteScreen tools_prosp, 13, 57
		EMWriteScreen tools_verif, 13, 69
		EMWriteScreen tools_imp_rel, 13, 75
		EMWriteScreen dues_retro, 14, 43
		EMWriteScreen dues_prosp, 14, 57
		EMWriteScreen dues_verif, 14, 69
		EMWriteScreen dues_imp_rel, 14, 75
		EMWriteScreen othr_retro, 15, 43
		EMWriteScreen othr_prosp, 15, 57
		EMWriteScreen othr_verif, 15, 69
		EMWriteScreen othr_imp_rel, 15, 75
	ELSE
		PF9
		'---If the script is editing an existing WKEX page, it would be doing so ONLY to update the HC Expense sub-menu.
		'---Adding to the HC Expenses
		EMWriteScreen "X", 18, 57
		transmit
		
		EMReadScreen current_month, 17, 20, 51
		IF current_month = "CURRENT MONTH + 1" THEN 
			PF3
		ELSE
			EMWriteScreen HC_Exp_Fed_Tax, 8, 36
			EMWriteScreen HC_Exp_State_Tax, 9, 36
			EMWriteScreen HC_Exp_FICA, 10, 36
			EMWriteScreen HC_Exp_Tran, 11, 36
			EMWriteScreen HC_Exp_Tran_imp_rel, 11, 51
			EMWriteScreen HC_Exp_Meals, 12, 36
			EMWriteScreen HC_Exp_Meals_Imp_Rel, 12, 51
			EMWriteScreen HC_Exp_Uniforms, 13, 36
			EMWriteScreen HC_Exp_Uniforms_Imp_Rel, 13, 51
			EMWriteScreen HC_Exp_Tools, 14, 36
			EMWriteScreen HC_Exp_Tools_Imp_Rel, 14, 51
			EMWriteScreen HC_Exp_Dues, 15, 36
			EMWriteScreen HC_Exp_Dues_Imp_Rel, 15, 51
			EMWriteScreen HC_Exp_Othr, 16, 36
			EMWriteScreen HC_Exp_Othr_Imp_Rel, 16, 51
			transmit
			PF3
		END IF
	END IF
	transmit
END FUNCTION

FUNCTION write_panel_to_MAXIS_WREG(wreg_fs_pwe, wreg_fset_status, wreg_defer_fs, wreg_fset_orientation_date, wreg_fset_sanction_date, wreg_num_sanctions, wreg_abawd_status, wreg_ga_basis)
	call navigate_to_screen("STAT", "WREG")
	call create_panel_if_nonexistent

	EMWriteScreen wreg_fs_pwe, 6, 68
	EMWriteScreen wreg_fset_status, 8, 50
	EMWriteScreen wreg_defer_fs, 8, 80
	IF wreg_fset_orientation_date <> "" THEN call create_MAXIS_friendly_date(wreg_fset_orientation_date, 0, 9, 50)
	IF wreg_fset_sanction_date <> "" THEN call create_MAXIS_friendly_date(wreg_fset_orientation_date, 0, 10, 50)
	IF wreg_num_sanctions <> "" THEN EMWriteScreen wreg_num_sanctions, 11, 50
	EMWriteScreen wreg_abawd_status, 13, 50
	EMWriteScreen wreg_ga_basis, 15, 50

	transmit
END FUNCTION
