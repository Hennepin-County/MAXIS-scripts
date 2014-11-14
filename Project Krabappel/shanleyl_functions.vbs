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

Function create_if_nonexistant()
	EMWriteScreen reference_number , 20, 76
	transmit
	EMReadScreen case_panel_check, 44, 24, 2
	If case_panel_check = "REFERENCE NUMBER IS NOT VALID FOR THIS PANEL" then
		EMReadScreen quantity_of_screens, 1, 2, 78
		If quantity_of_screens <> "0" then
			PF9
		ElseIf quantity_of_screens = "0" then
			EMWriteScreen "__", 20, 76
			EMWriteScreen "NN", 20, 79
			Transmit
		End If
	ElseIf case_panel_check <> "REFERENCE NUMBER IS NOT VALID FOR THIS PANEL" then
		EMReadScreen error_scan, 80, 24, 1
		error_scan = trim(error_scan)
		EMReadScreen quantity_of_screens, 1, 2, 78
		If error_scan = "" and quantity_of_screens <> "0" then
			PF9
		ElseIf error_scan <> "" and quantity_of_screens <> "0" then
			'FIX ERROR HERE
			msgbox("Error: " & error_scan)
		ElseIf error_scan <> "" and quantity_of_screens = "0" then
			EMWriteScreen reference_number, 20, 76
			EMWriteScreen "NN", 20, 79
			Transmit
		End If
	End If
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

'--------------- PANEL FUNCTIONS ---------------------

Function write_panel_to_maxis_MSUR(msur_begin_date)
	call maxis_dater(msur_begin_date,return_date,"MNSure Begin Date")
	call navigate_to_screen("STAT","MSUR")
	call create_if_nonexistant
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

Function write_panel_to_maxis_PDED(pded_wid_deduction,pded_adult_child_disregard,pded_wid_disregard,pded_earned_income_deduction_reason,pded_earned_income_deduction_value)
	call navigate_to_screen("STAT","PDED")
	call create_if_nonexistant
	'Pickle Disregard
		'ADD ME LATER
	'Disa Widow/ers Deduction
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
	
	'Blind/Disa Student Child Disregard
	
	'Guardianship Fee
	
	'Rep Payee Fee
	
	'Other Expense
	
	'Shelter Special Needs
	
	'Excess Need
	
	'Restaurant Meals
	
End Function

Function write_panel_to_maxis_HCRE()

End Function

Function write_panel_to_maxis_INSA()

End Function

Function write_panel_to_maxis_BUSI()

End Function

Function write_panel_to_maxis_DFLN()

End Function

Function write_panel_to_maxis_STWK()

End Function

Function write_panel_to_maxis_ABPS()

End Function

Function write_panel_to_maxis_WKEX()

End Function

Function write_panel_to_maxis_FMED()

End Function

'----------- TEST AREA -------------

case_number = "211486"
reference_number = "01"

'write_panel_to_maxis_MSUR(Any_Date)

EMConnect ""

'call write_panel_to_maxis_PDED()

call navigate_to_screen("STAT","EATS")
call create_if_nonexistant()

stopscript