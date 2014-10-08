Public Sub standard_date(A,B,C,D)
  'A = appl_month
  'B = appl_day
  'C = appl_year
  'D = Var Name to return
  If Len(B) = 1 then B = "0" & B
  If Len(A) = 1 then A = "0" & A
  If Len(C) = 4 then C = right(C,2)
  D = A & "/" & B & "/"&C	
End Sub

Public Sub get_time(A)
	A = replace(Mid(now,10,5), " ", "") & " " & Right(Now, 2)
End Sub

Function budget_month_config(A,B)
	'A = Budget Month
	'B = Return Variable
	'X = Offset
	'Y = Method Var
	'Z = Eligibility Var
	If A = 1 then
		X = 0			
	ElseIf A = 2 then
		X = 11			
	ElseIf A = 3 then
		X = 22			
	ElseIf A = 4 then
		X = 33			
	ElseIf A = 5 then
		X = 44			
	ElseIf A = 6 then
		X = 55			
	End If		
	EMReadScreen Y, 1, 13, 76
	EMWriteScreen Y, 13, 21 + X
  If Y = "B" then
    EMReadScreen B_elig_type, 2, 12, 72
      EMWriteScreen B_elig_type, 12, 17 + X
    EMReadScreen B_max_standard, 1, 12, 77
    EMWriteScreen B_max_standard, 12, 22 + X
  ElseIf Y = "A" then
    EMReadScreen Z, 2, 12, 72
      EMWriteScreen Z, 12, 17 + X
    EMReadScreen B, 5, 6, 19 + X
    If B <> "" then call elig_to_standard(Right(budget_one,2), Z, maxis_standard)
    EMWriteScreen maxis_standard, 12, 22 + X
  End If
End Function

Public Sub retro_calculator(A,B,C,D)
  'A = Number of retro months requested
  'B = Number of gap months needed
  'C = Input date
  'D = Name of return variable
  D = DateAdd("m", -A-B,C)
  retro_month = Replace(Left(DateAdd("m", -A-B,C),2),"/","")
  If retro_month < 10 then retro_month = "0"&retro_month
  retro_year	= Right(DateAdd("yyyy", - DateDiff("yyyy", D, C), C), 2)
  D = retro_month & "/" & retro_year
End Sub

Public Function add_days(D,E,F)
  'D = Days to Add or Subtract 
  'E = Starting Date
  'F = Var to name the return variable
  calc_date = DateAdd("d", D, E)
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

Public Function add_months(D,E,F)
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

function dail_writ(A,B)
  'A = Days to add
  'B = Dail Message
  navigate_to_screen "DAIL",""
  EMWriteScreen maxis_case_number,20,38
  EMWriteScreen "WRIT",20,70
  Transmit
  call add_days(A,appl_date,post_calc_date)
  EMWriteScreen post_calc_date,5,18
  EMSetCursor 9,3
  EMSendKey (B)
End Function

Public Function maxis_date(D,E)
	'A = Month
	'B = Day
	'C = Year
	'D = Input Date
	'E = Return Variable
	'F = Conversion Variable
  F = D
  F = trim(F)
  F = replace(F, "-", " ")
  F = replace(F, "/", " ")
  F = replace(F, "\", " ")
  F = replace(F, " ", "     ")
  A = left(F, 5)
  A = trim(A)
  If Len(A) = 1 then A = "0" & A
  B = Mid(F, 6, 4)
  B = trim(B)
  If Len(B) = 1 then B = "0" & B
  C = Right(F, 2)
  C = trim(C)
  E = A & "/" & B & "/" & C
End Function

Public Function maxis_dater(A,B,C)
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

Function current_date(A)
	'A = Name of return variable
	maxis_date Date,A
End Function

function run_file(file_path)
  Set run_file_fso = CreateObject("Scripting.FileSystemObject")
  Set fso_run_file_command = run_file_fso.OpenTextFile(file_path)
  file_contents = fso_run_file_command.ReadAll
  fso_run_file_command.Close
  Execute file_contents
end function

function case_in_background_delay
  call navigate_to_screen("STAT","SUMM")
    Do
      EMReadScreen bgtx_popup, 10, 7, 32
	  bgtx_popup = trim(bgtx_popup)
	  bgtx_popup = replace(bgtx_popup, ".", "")
      EMReadScreen bgtx_locked, 12, 24, 27
	  bgtx_locked = trim(bgtx_locked)
	  bgtx_locked = replace(bgtx_locked, ".", "")
	  If bgtx_popup = "Background" then
        EMWriteScreen "N", 12, 47
        transmit
        EMWriteScreen "WAIT", 16, 43
        EMWaitReady 10, 2000
        call navigate_to_screen("STAT","SUMM")
      ElseIf bgtx_locked = "BACKGROUND" then
        EMWriteScreen "WAIT", 16, 43
        EMWaitReady 10, 2000
        call navigate_to_screen("STAT","SUMM")
      End If
    Loop until bgtx_popup <> "Background" and bgtx_locked <> "BACKGROUND"
end function

function hhmm_background_check
	Do
	EMReadScreen hhmm_background, 10, 24, 25
		If hhmm_background = "BACKGROUND" then
			EMWriteScreen "WAIT", 20, 71
			EMWaitReady 10, 2000
			call change_footer_month(retro_month_requested,retro_year_requested)
		End If
	Loop until hhmm_background <> "BACKGROUND"
end function

function send_case_through_background(update_check)
  if update_check = "no_update" then
    update_check = trim(update_check)
	update_check = "N"
    call navigate_to_screen("STAT","BGTX")
    EMWriteScreen update_check, 16, 54
    transmit
	call case_in_background_delay
  ElseIf update_check = "update" then
    update_check = trim(update_check)
	update_check = "Y"
    call navigate_to_screen("STAT","BGTX")
    EMWriteScreen update_check, 16, 54
    transmit
  End If
end function

function stat_error_scanner  
  BeginDialog summ_error_review_dialog, 0, 0, 170, 63, "Please review stat edits"
    ButtonGroup ButtonPressed
      OkButton 114, 49, 22, 12
      CancelButton 139, 49, 30, 12
    Text 2, 3, 168, 41, "Leave this screen up while your correct the STAT edits. Please review the STAT edits. If any of them pertain to health care please resolve them before continuing. Once completed select OK to continue or select CANCEL to stop running the script."
  EndDialog
  EMReadScreen on_self_menu_check, 4, 2, 50
  on_self_menu_check = trim(on_self_menu_check)
  If on_self_menu_check <> "SELF" then
    EMWriteScreen "SELF", 20, 71
    transmit
  End If
  EMWriteScreen appl_month, 20, 43
  EMWriteScreen appl_year, 20, 46
  call case_in_background_delay
  EMReadScreen stat_errors_exist, 4, 5, 3
  stat_errors_exist = trim(stat_errors_exist)
  IF stat_errors_exist <> "" then
  Dialog summ_error_review_dialog
    If buttonpressed = 0 then stopscript
  End If
  call send_case_through_background("no_update")
  call case_in_background_delay
end function

function command_line(first_line,middle_line,last_line)
  first_line = trim(first_line)
  middle_line = trim(middle_line)
  last_line = trim(last_line)
  If first_line = "ignore" then first_line = "____"
  If middle_line = "ignore" then middle_line = "__"
  If last_line = "ignore" then last_line = "__"
  row = 1
  col = 1
  EMSearch "Command: ", row, col
  If row <> 0 then EMWriteScreen first_line, row, col + 9
  If row <> 0 then EMWriteScreen middle_line, row, col + 14 
  If row <> 0 then EMWriteScreen last_line, row, col + 17
  transmit
end function

function change_footer_month(new_footer_month,new_footer_year)
  new_footer_month = trim(new_footer_month)
  new_footer_year = trim(new_footer_year)
  EMReadScreen on_self_menu_check, 4, 2, 50
  on_self_menu_check = trim(on_self_menu_check)  
  If on_self_menu_check = "SELF" then
    EMWriteScreen new_footer_month, 20, 43
    EMWriteScreen new_footer_year, 20, 46
  ElseIf on_self_menu_check <> "SELF" then
    row = 1
    col = 1
    EMSearch "Month: ", row, col
    If row <> 0 then EMWriteScreen new_footer_month, row, col + 7
    If row <> 0 then EMWriteScreen new_footer_year, row, col + 10
    transmit
  End If
end function

function elig_to_standard(A,B,C)
  'A = year
  'B = Eligibility Type
  'C = Return Variable
  'D = Error Msg
  'E = Error Box Title
  D = "An error has occurred and the current eligibility type cannot be converted at this time. This is likely if the 20" & A & " eligibility types have not been programmed into the script. Please set this manually before closing this window. Please be aware this message will display every time this error occurs."
  E = "Conversion error"
  If A = "13" then
    If B = "AX" then
      C = " "
	  MsgBox D,E
    ElseIf B = "AA" then
      C = " "
	  MsgBox D,E
	ElseIf B = "CX" then
      C = " "
	  MsgBox D,E
    ElseIf B = "PX" then
	  C = " "
	  MsgBox D,E
	ElseIf B = "PC" then
      C = " "
	  MsgBox D,E
    ElseIf B = "CK" then
      C = " "
	  MsgBox D,E
    ElseIf B = "CB" then
      C = " "
	  MsgBox D,E
	End If
  ElseIf A = "14" then
    If B = "AX" or B = "AA" or B = "CX" then
      C = "J"
    ElseIf B = "PX" or B = "PC" then
      C = "T"
    ElseIf B = "CK" then
      C = "K"
    ElseIf B = "CB" then
      C = "I"
	End If
  End If
end function

Function elig_member_selection()
  client_01_exists_in_list = false
  client_02_exists_in_list = false
  client_03_exists_in_list = false
  client_04_exists_in_list = false
  client_05_exists_in_list = false
  client_06_exists_in_list = false
  client_07_exists_in_list = false
  client_08_exists_in_list = false
  client_09_exists_in_list = false
  client_10_exists_in_list = false
  client_11_exists_in_list = false
  client_12_exists_in_list = false   
  extra_questions_01 = true
  extra_questions_02 = true 
  extra_questions_03 = true
  extra_questions_04 = true
  extra_questions_05 = true
  extra_questions_06 = true
  extra_questions_07 = true
  extra_questions_08 = true
  extra_questions_09 = true
  extra_questions_10 = true
  extra_questions_11 = true
  extra_questions_12 = true   
  current_client_selected = ""
  EMReadScreen HH_count_01, 25, 12, 13
  EMReadScreen HH_count_02, 25, 13, 13
  EMReadScreen HH_count_03, 25, 14, 13
  EMReadScreen HH_count_04, 25, 15, 13
  EMReadScreen HH_count_05, 25, 16, 13
  EMReadScreen HH_count_06, 25, 17, 13
  EMReadScreen HH_count_07, 25, 18, 13
  EMReadScreen HH_count_08, 25, 19, 13
  EMReadScreen HH_count_09, 25, 20, 13
  EMReadScreen HH_count_10, 25, 21, 13
  EMReadScreen HH_count_11, 25, 22, 13
  EMReadScreen HH_count_12, 25, 23, 13
  set_dialog_size = 82
  set_button_position = 65
  set_window_width = 260
  
  EMReadScreen current_client_selected, 25, 5, 14
	current_client_selected = trim(current_client_selected) 
  
  If HH_count_01 <> "                         " then 
    set_dialog_size = 82
    set_button_position = 65
	set_window_width = 260
	client_01_exists_in_list = true
	HH_count_01 = trim(HH_count_01)
	hh_count_01_searchable = replace(HH_count_01, "  ", " ")
  End If
  If HH_count_02 <> "                         " then 
    set_dialog_size = 122
    set_button_position = 105
	set_window_width = 260
	client_02_exists_in_list = true
	HH_count_02 = trim(HH_count_02)
	hh_count_02_searchable = replace(HH_count_02, "  ", " ")
  End If
  If HH_count_03 <> "                         " then 
    set_dialog_size = 162
    set_button_position = 145	
	set_window_width = 260
	client_03_exists_in_list = true
	HH_count_03 = trim(HH_count_03)
	hh_count_03_searchable = replace(HH_count_03, "  ", " ")
  End If
  If HH_count_04 <> "                         " then 
    set_dialog_size = 202
    set_button_position = 185	
	set_window_width = 260
	client_04_exists_in_list = true
	HH_count_04 = trim(HH_count_04)
	hh_count_04_searchable = replace(HH_count_04, "  ", " ")
  End If
  If HH_count_05 <> "                         " then 
    set_dialog_size = 242
    set_button_position = 225	
	set_window_width = 260
	client_05_exists_in_list = true
	HH_count_05 = trim(HH_count_05)
	hh_count_05_searchable = replace(HH_count_05, "  ", " ")
  End If
  If HH_count_06 <> "                         " then 
    set_dialog_size = 282
    set_button_position = 265
	set_window_width = 260
	client_06_exists_in_list = true
	HH_count_06 = trim(HH_count_06)
	hh_count_06_searchable = replace(HH_count_06, "  ", " ")
  End If
  If HH_count_07 <> "                         " then 
    set_dialog_size = 282
    set_button_position = 265
	set_window_width = 515
	client_07_exists_in_list = true
	HH_count_07 = trim(HH_count_07)
	hh_count_07_searchable = replace(HH_count_07, "  ", " ")
  End If
  If HH_count_08 <> "                         " then 
    set_dialog_size = 282
    set_button_position = 265
	set_window_width = 515
	client_08_exists_in_list = true
	HH_count_08 = trim(HH_count_08)
	hh_count_08_searchable = replace(HH_count_08, "  ", " ")
  End If
  If HH_count_09 <> "                         " then 
    set_dialog_size = 282
	set_button_position = 265
	set_window_width = 515
	client_09_exists_in_list = true
	HH_count_09 = trim(HH_count_09)
	hh_count_09_searchable = replace(HH_count_09, "  ", " ")
  End If
  If HH_count_10 <> "                         " then 
    set_dialog_size = 282
	set_button_position = 265
	set_window_width = 515
	client_10_exists_in_list = true
	HH_count_10 = trim(HH_count_10)
	hh_count_10_searchable = replace(HH_count_10, "  ", " ")
  End If
  If HH_count_11 <> "                         " then 
    set_dialog_size = 282
	set_button_position = 265
	set_window_width = 515
	client_11_exists_in_list = true
	HH_count_11 = trim(HH_count_11)
	hh_count_11_searchable = replace(HH_count_11, "  ", " ")
  End If
  If HH_count_12 <> "                         " then 
    set_dialog_size = 282
	set_button_position = 265
	set_window_width = 515
	client_12_exists_in_list = true
	HH_count_12 = trim(HH_count_12)
	hh_count_12_searchable = replace(HH_count_12, "  ", " ")
  End If

  If HH_count_01 = current_client_selected then extra_questions_01 = false
  If HH_count_02 = current_client_selected then extra_questions_02 = false 
  If HH_count_03 = current_client_selected then extra_questions_03 = false
  If HH_count_04 = current_client_selected then extra_questions_04 = false
  If HH_count_05 = current_client_selected then extra_questions_05 = false
  If HH_count_06 = current_client_selected then extra_questions_06 = false
  If HH_count_07 = current_client_selected then extra_questions_07 = false
  If HH_count_08 = current_client_selected then extra_questions_08 = false
  If HH_count_09 = current_client_selected then extra_questions_09 = false
  If HH_count_10 = current_client_selected then extra_questions_10 = false
  If HH_count_11 = current_client_selected then extra_questions_11 = false
  If HH_count_12 = current_client_selected then extra_questions_12 = false
  
  BeginDialog elig_HH_count_dialog, 0, 0, set_window_width, set_dialog_size, "Budget Configuration"
    ButtonGroup ButtonPressed
      OkButton 200, set_button_position, 20, 12
      CancelButton 225, set_button_position, 30, 12
    Text 60, 8, 190, 10, "Elig Configuration for "&current_client_selected
	
    If client_01_exists_in_list = true then 
	  GroupBox 5, 20, 250, 40, HH_count_01
      Text 10, 40,  80, 8, "Select type(s) of income:"	
	  If extra_questions_01 = false then
	    client_01_hh_count = "Yes"
	    client_01_include_income = "Yes"
      ElseIf extra_questions_01 = true then
        Text 10, 50, 210, 8, "Does income deem to "&current_client_selected &"?"
        Text 10, 30, 210, 8, "Is this person included in "&current_client_selected &"'s household?"
	    DropListBox 220, 26, 30, 10, "Yes"+chr(9)+"No", client_01_hh_count
        DropListBox 220, 45, 30, 10, ""+chr(9)+"Yes"+chr(9)+"No", client_01_include_income
	  End If
	  CheckBox 92, 40, 58, 8, "Earned Income", earned_income_01
	  CheckBox 151, 40, 67, 8, "Unearned Income", unearned_income_01	  
	End If  
	
    If client_02_exists_in_list = true then 
	  GroupBox 5, 60, 250, 40, HH_count_02
	  Text 10, 80, 80,  8, "Select type(s) of income:"
	  If extra_questions_02 = false then
	    client_02_hh_count = "Yes"
	    client_02_include_income = "Yes"
	  ElseIf extra_questions_02 = true then
	    Text 10, 70, 210, 8, "Is this person included in "&current_client_selected &"'s household?"
  	    Text 10, 90, 210, 8, "Does income deem to "&current_client_selected &"?"
        DropListBox 220, 66, 30, 10, "Yes"+chr(9)+"No", client_02_hh_count
	    DropListBox 220, 85, 30, 10, ""+chr(9)+"Yes"+chr(9)+"No", client_02_include_income
	  End If
	  CheckBox 92, 80, 58, 8, "Earned Income", earned_income_02
	  CheckBox 151, 80, 67, 8, "Unearned Income", unearned_income_02 
	End If
	
    If client_03_exists_in_list = true then 
	  GroupBox 5, 100, 250, 40, HH_count_03
	  Text 10, 120,  80, 8, "Select type(s) of income:"
      If extra_questions_03 = false then
	    client_03_hh_count = "Yes"
	    client_03_include_income = "Yes"
  	  ElseIf extra_questions_03 = true then
  	    Text 10, 130, 210, 8, "Does income deem to "&current_client_selected &"?"
  	    Text 10, 110, 210, 8, "Is this person included in "&current_client_selected &"'s household?"
        DropListBox 220, 106, 30, 10, "Yes"+chr(9)+"No", client_03_hh_count	  
        DropListBox 220, 125, 30, 10, ""+chr(9)+"Yes"+chr(9)+"No", client_03_include_income
	  End If
	  CheckBox 92, 120, 58, 8, "Earned Income", earned_income_03
	  CheckBox 151, 120, 67, 8, "Unearned Income", unearned_income_03	  
	End If
	
    If client_04_exists_in_list = true then 
	  GroupBox 5, 140, 250, 40, HH_count_04
	  Text 10, 160,  80, 8, "Select type(s) of income:"
	  IF extra_questions_04 = false then
	    client_04_hh_count = "Yes"
	    client_04_include_income = "Yes"
	  ElseIf extra_questions_04 = true then
	  	Text 10, 170, 210, 8, "Does income deem to "&current_client_selected &"?"
	  	Text 10, 150, 210, 8, "Is this person included in "&current_client_selected &"'s household?"
        DropListBox 220, 146, 30, 10, "Yes"+chr(9)+"No", client_04_hh_count
	    DropListBox 220, 165, 30, 10, ""+chr(9)+"Yes"+chr(9)+"No", client_04_include_income
	  End If
	  CheckBox 92, 160, 58, 8, "Earned Income", earned_income_04
	  CheckBox 151, 160, 67, 8, "Unearned Income", unearned_income_04	  
	End If
	
    If client_05_exists_in_list = true then 
	  GroupBox 5, 180, 250, 40, HH_count_05
	  Text 10, 200,  80, 8, "Select type(s) of income:"
	  If extra_questions_05 = false then
	    client_05_hh_count = "Yes"
	    client_05_include_income = "Yes"
	  ElseIf extra_questions_05 = true then
  	    Text 10, 190, 210, 8, "Is this person included in "&current_client_selected &"'s household?"
  	    Text 10, 210, 210, 8, "Does income deem to "&current_client_selected &"?"
	    DropListBox 220, 186, 30, 10, "Yes"+chr(9)+"No", client_05_hh_count
	    DropListBox 220, 205, 30, 10, ""+chr(9)+"Yes"+chr(9)+"No", client_05_include_income
	  End If
	  CheckBox 92, 200, 58, 8, "Earned Income", earned_income_05
	  CheckBox 151, 200, 67, 8, "Unearned Income", unearned_income_05	  
	End If
	
    If client_06_exists_in_list = true then 
	  GroupBox 5, 220, 250, 40, HH_count_06
	  Text 10, 240,  80, 8, "Select type(s) of income:"
	  If extra_questions_06 = false then
	    client_06_hh_count = "Yes"
	    client_06_include_income = "Yes"
	  ElseIf extra_questions_06 = true then
        Text 10, 230, 210, 8, "Is this person included in "&current_client_selected &"'s household?"
  	    Text 10, 250, 210, 8, "Does income deem to "&current_client_selected &"?"
	    DropListBox 220, 226, 30, 10, "Yes"+chr(9)+"No", client_06_hh_count
	    DropListBox 220, 245, 30, 10, ""+chr(9)+"Yes"+chr(9)+"No", client_06_include_income
	  End If
	  CheckBox 92, 240, 58, 8, "Earned Income", earned_income_06
	  CheckBox 157, 240, 67, 8, "Unearned Income", unearned_income_06
	End If
	
    If client_07_exists_in_list = true then
      GroupBox 260, 20, 250, 40, HH_count_07	  
	  Text 265, 40,  80, 8, "Select type(s) of income:"	
	  If extra_questions_07 = false then
	    client_07_hh_count = "Yes"
	    client_07_include_income = "Yes"
	  ElseIf extra_questions_07 = true then
        Text 265, 30, 210, 8, "Is this person included in "&current_client_selected &"'s household?"
        Text 265, 50, 210, 8, "Does income deem to "&current_client_selected &"?"	
	    DropListBox 475, 26, 30, 10, "Yes"+chr(9)+"No", client_07_hh_count
	    DropListBox 475, 45, 30, 10, ""+chr(9)+"Yes"+chr(9)+"No", client_07_include_income
	  End If
	  CheckBox 347, 40, 58, 8, "Earned Income", earned_income_07
	  CheckBox 406, 40, 67, 8, "Unearned Income", unearned_income_07
	End If
	
    If client_08_exists_in_list = true then
      GroupBox 260, 60, 250, 40, HH_count_08	  
	  Text 265, 80,  80, 8, "Select type(s) of income:"	
      If extra_questions_08 = false then
	    client_08_hh_count = "Yes"
	    client_08_include_income = "Yes"
	  ElseIf extra_questions_08 = true then
        Text 265, 70, 210, 8, "Is this person included in "&current_client_selected &"'s household?"
        Text 265, 90, 210, 8, "Does income deem to "&current_client_selected &"?"	
        DropListBox 475, 66, 30, 10, "Yes"+chr(9)+"No", client_08_hh_count
        DropListBox 475, 85, 30, 10, ""+chr(9)+"Yes"+chr(9)+"No", client_08_include_income
	  End If		
      CheckBox 347, 80, 58, 8, "Earned Income", earned_income_08
      CheckBox 406, 80, 67, 8, "Unearned Income", unearned_income_08  
	End If
	
    If client_09_exists_in_list = true then
      GroupBox 260, 100, 250, 40, HH_count_09
	  Text 265, 120, 80, 8, "Select type(s) of income:"
	  If extra_questions_09 = false then
	    client_09_hh_count = "Yes"
	    client_09_include_income = "Yes"
      ElseIf extra_questions_09 = true then
        Text 265, 110, 210, 8, "Is this person included in "&current_client_selected &"'s household?" 
	    Text 265, 130, 210, 8, "Does income deem to "&current_client_selected &"?"	
	    DropListBox 475, 106, 30, 10, "Yes"+chr(9)+"No", client_09_hh_count
        DropListBox 475, 125, 30, 10, ""+chr(9)+"Yes"+chr(9)+"No", client_09_include_income	
	  End If
      CheckBox 347, 120, 58, 8, "Earned Income", earned_income_09
      CheckBox 406, 120, 67, 8, "Unearned Income", unearned_income_09  
	End If
	
    If client_10_exists_in_list = true then 
      GroupBox 260, 140, 250, 40, HH_count_10
	  Text 265, 160,  80, 8, "Select type(s) of income:"
	  If extra_questions_10 = false then
	    client_10_hh_count = "Yes"
	    client_10_include_income = "Yes"
      ElseIf extra_questions_10 = true then
        Text 265, 150, 210, 8, "Is this person included in "&current_client_selected &"'s household?"
	    Text 265, 170, 210, 8, "Does income deem to "&current_client_selected &"?"
	    DropListBox 475, 146, 30, 10, "Yes"+chr(9)+"No", client_10_hh_count
	    DropListBox 475, 165, 30, 10, ""+chr(9)+"Yes"+chr(9)+"No", client_10_include_income
	  End If
	  CheckBox 347, 160, 58, 8, "Earned Income", earned_income_10
	  CheckBox 406, 160, 67, 8, "Unearned Income", unearned_income_10	  
	End If
	
    If client_11_exists_in_list = true then
      GroupBox 260, 180, 250, 40, HH_count_11
      Text 265, 200,  80, 8, "Select type(s) of income:"	
	  If extra_questions_11 = false then
	    client_11_hh_count = "Yes"
	    client_11_include_income = "Yes"
	  ElseIf extra_questions_11 = true then
        Text 265, 190, 210, 8, "Is this person included in "&current_client_selected &"'s household?"	
        Text 265, 210, 210, 8, "Does income deem to "&current_client_selected &"?"
	    DropListBox 475, 186, 30, 10, "Yes"+chr(9)+"No", client_11_hh_count
	    DropListBox 475, 205, 30, 10, ""+chr(9)+"Yes"+chr(9)+"No", client_11_include_income
	  End If
	  CheckBox 347, 200, 58, 8, "Earned Income", earned_income_11
	  CheckBox 406, 200, 67, 8, "Unearned Income", unearned_income_11
	End If
	
    If client_12_exists_in_list = true then 
      GroupBox 260, 220, 250, 40, HH_count_12
	  Text 265, 240,  80, 8, "Select type(s) of income:"
	  If extra_questions_12 = false then
	    client_12_hh_count = "Yes"
	    client_12_include_income = "Yes"
	  ElseIf extra_questions_12 = true then
	  	Text 265, 230, 210, 8, "Is this person included in "&current_client_selected &"'s household?"
	  	Text 265, 250, 210, 8, "Does income deem to "&current_client_selected &"?"
	    DropListBox 475, 226, 30, 10, "Yes"+chr(9)+"No", client_12_hh_count
	    DropListBox 475, 245, 30, 10, ""+chr(9)+"Yes"+chr(9)+"No", client_12_include_income
	  End If
	  CheckBox 347, 240, 58, 8, "Earned Income", earned_income_12
	  CheckBox 406, 240, 67, 8, "Unearned Income", unearned_income_12  
	End If
	
  EndDialog
  
  Dialog elig_HH_count_dialog
    If buttonpressed = 0 then stopscript
  
  If client_01_hh_count = "Yes" then 
    client_01_hh_count = 1
  ElseIf client_01_hh_count <> "Yes" or client_01_exists_in_list = false then 
    client_01_hh_count = 0
  End If
  If client_02_hh_count = "Yes" then 
    client_02_hh_count = 1
  ElseIf client_02_hh_count <> "Yes" or client_02_exists_in_list = false then 
    client_02_hh_count = 0
  End If
  If client_03_hh_count = "Yes" then 
    client_03_hh_count = 1
  ElseIf client_03_hh_count <> "Yes" or client_03_exists_in_list = false then 
    client_03_hh_count = 0
  End If
  If client_04_hh_count = "Yes" then 
    client_04_hh_count = 1
  ElseIf client_04_hh_count <> "Yes" or client_04_exists_in_list = false then 
    client_04_hh_count = 0
  End If
  If client_05_hh_count = "Yes" then 
    client_05_hh_count = 1
  ElseIf client_05_hh_count <> "Yes" or client_05_exists_in_list = false then 
    client_05_hh_count = 0
  End If
  If client_06_hh_count = "Yes" then 
    client_06_hh_count = 1
  ElseIf client_06_hh_count <> "Yes" or client_06_exists_in_list = false then 
    client_06_hh_count = 0
  End If
  If client_07_hh_count = "Yes" then 
    client_07_hh_count = 1
  ElseIf client_07_hh_count <> "Yes" or client_07_exists_in_list = false then 
    client_07_hh_count = 0
  End If
  If client_08_hh_count = "Yes" then 
    client_08_hh_count = 1
  ElseIf client_08_hh_count <> "Yes" or client_08_exists_in_list = false then 
    client_08_hh_count = 0
  End If
  If client_09_hh_count = "Yes" then 
    client_09_hh_count = 1
  ElseIf client_09_hh_count <> "Yes" or client_09_exists_in_list = false then 
    client_09_hh_count = 0
  End If
  If client_10_hh_count = "Yes" then 
    client_10_hh_count = 1
  ElseIf client_10_hh_count <> "Yes" or client_10_exists_in_list = false then 
    client_10_hh_count = 0
  End If
  If client_11_hh_count = "Yes" then 
    client_11_hh_count = 1
  ElseIf client_11_hh_count <> "Yes" or client_11_exists_in_list = false then 
    client_11_hh_count = 0
  End If
  If client_12_hh_count = "Yes" then 
    client_12_hh_count = 1
  ElseIf client_12_hh_count <> "Yes" or client_12_exists_in_list = false then 
    client_12_hh_count = 0
  End If
   
  total_hh_count = client_01_hh_count + client_02_hh_count + client_03_hh_count + client_04_hh_count + client_05_hh_count + client_06_hh_count + client_07_hh_count + client_08_hh_count + client_09_hh_count + client_10_hh_count + client_11_hh_count + client_12_hh_count
  total_hh_count = trim(total_hh_count)
  
  If client_01_include_income = "Yes" then 
    set_included_income_01 = "Y"
  ElseIf client_01_include_income <> "Yes" and client_01_exists_in_list = true then 
    set_included_income_01 = "N"
  End If
  If client_02_include_income = "Yes" then 
    set_included_income_02 = "Y"
  ElseIf client_02_include_income <> "Yes" and client_02_exists_in_list = true then 
    set_included_income_02 = "N"
  End If
  If client_03_include_income = "Yes" then 
    set_included_income_03 = "Y"
  ElseIf client_03_include_income <> "Yes" and client_03_exists_in_list = true then 
    set_included_income_03 = "N"
  End If
  If client_04_include_income = "Yes" then 
    set_included_income_04 = "Y"
  ElseIf client_04_include_income <> "Yes" and client_04_exists_in_list = true then 
    set_included_income_04 = "N"
  End If
  If client_05_include_income = "Yes" then 
    set_included_income_05 = "Y"
  ElseIf client_05_include_income <> "Yes" and client_05_exists_in_list = true then 
    set_included_income_05 = "N"
  End If
  If client_06_include_income = "Yes" then 
    set_included_income_06 = "Y"
  ElseIf client_06_include_income <> "Yes" and client_06_exists_in_list = true then 
    set_included_income_06 = "N"
  End If
  If client_07_include_income = "Yes" then 
    set_included_income_07 = "Y"
  ElseIf client_07_include_income <> "Yes" and client_07_exists_in_list = true then 
    set_included_income_07 = "N"
  End If
  If client_08_include_income = "Yes" then 
    set_included_income_08 = "Y"
  ElseIf client_08_include_income <> "Yes" and client_08_exists_in_list = true then 
    set_included_income_08 = "N"
  End If
  If client_09_include_income = "Yes" then 
    set_included_income_09 = "Y"
  ElseIf client_09_include_income <> "Yes" and client_09_exists_in_list = true then 
    set_included_income_09 = "N"
  End If
  If client_10_include_income = "Yes" then 
    set_included_income_10 = "Y"
  ElseIf client_10_include_income <> "Yes" and client_10_exists_in_list = true then 
    set_included_income_10 = "N"
  End If
  If client_11_include_income = "Yes" then 
    set_included_income_11 = "Y"
  ElseIf client_11_include_income <> "Yes" and client_11_exists_in_list = true then 
    set_included_income_11 = "N"
  End If
  If client_12_include_income = "Yes" then 
    set_included_income_12 = "Y"
  ElseIf client_12_include_income <> "Yes" and client_12_exists_in_list = true then 
    set_included_income_12 = "N"
  End If
  
  If client_01_exists_in_list = false then earned_income_01 = ""
  If client_02_exists_in_list = false then earned_income_02 = ""
  If client_03_exists_in_list = false then earned_income_03 = ""
  If client_04_exists_in_list = false then earned_income_04 = ""
  If client_05_exists_in_list = false then earned_income_05 = ""
  If client_06_exists_in_list = false then earned_income_06 = ""
  If client_07_exists_in_list = false then earned_income_07 = ""
  If client_08_exists_in_list = false then earned_income_08 = ""
  If client_09_exists_in_list = false then earned_income_09 = ""
  If client_10_exists_in_list = false then earned_income_10 = ""
  If client_11_exists_in_list = false then earned_income_11 = ""
  If client_12_exists_in_list = false then earned_income_12 = ""
  
  If client_01_exists_in_list = false then unearned_income_01 = ""
  If client_02_exists_in_list = false then unearned_income_02 = ""
  If client_03_exists_in_list = false then unearned_income_03 = ""
  If client_04_exists_in_list = false then unearned_income_04 = ""
  If client_05_exists_in_list = false then unearned_income_05 = ""
  If client_06_exists_in_list = false then unearned_income_06 = ""
  If client_07_exists_in_list = false then unearned_income_07 = ""
  If client_08_exists_in_list = false then unearned_income_08 = ""
  If client_09_exists_in_list = false then unearned_income_09 = ""
  If client_10_exists_in_list = false then unearned_income_10 = ""
  If client_11_exists_in_list = false then unearned_income_11 = ""
  If client_12_exists_in_list = false then unearned_income_12 = ""
  
  If client_01_exists_in_list = true then client_01_name = HH_count_01
  If client_02_exists_in_list = true then client_02_name = HH_count_02
  If client_03_exists_in_list = true then client_03_name = HH_count_03
  If client_04_exists_in_list = true then client_04_name = HH_count_04
  If client_05_exists_in_list = true then client_05_name = HH_count_05
  If client_06_exists_in_list = true then client_06_name = HH_count_06
  If client_07_exists_in_list = true then client_07_name = HH_count_07
  If client_08_exists_in_list = true then client_08_name = HH_count_08
  If client_09_exists_in_list = true then client_09_name = HH_count_09
  If client_10_exists_in_list = true then client_10_name = HH_count_10
  If client_11_exists_in_list = true then client_11_name = HH_count_11
  If client_12_exists_in_list = true then client_12_name = HH_count_12
 
End function

Function additional_income_info(client_name,unearned_income_request,earned_income_request)
client_name = trim(client_name)

'-- Grab Budget Month and year ----------------

EMReadScreen elig_abud_check, 4, 3, 47
	elig_abud_check = trim(elig_abud_check)
EMReadScreen elig_cbud_check, 4, 3, 54
	elig_cbud_check = trim(elig_cbud_check)
If elig_abud_check = "ABUD" then
	EMReadScreen grab_budget_month, 2, 6, 11
		grab_budget_month = trim(grab_budget_month)
	EMReadScreen grab_budget_year, 2, 6, 14
		grab_budget_year = trim(grab_budget_year)
ElseIf elig_cbud_check = "CBUD" then
	EMReadScreen grab_budget_month, 2, 6, 14
		grab_budget_month = trim(grab_budget_month)
	EMReadScreen grab_budget_year, 2, 6, 17
		grab_budget_year = trim(grab_budget_year)
End If

'----------------------------------------------

'Rinsing the variables------
	  earned_type_1 = ""
	  earned_type_2 = ""
	  earned_type_3 = ""
      unearned_type_1 = ""
	  unearned_type_2 = ""
	  unearned_type_3 = ""
	  earned_value_1 = ""
	  earned_value_2 = ""
	  earned_value_3 = ""
      unearned_value_1 = ""
	  unearned_value_2 = ""
	  unearned_value_3 = ""
      earned_exclusion_1 = ""
	  earned_exclusion_2 = ""
	  earned_exclusion_3 = ""
      unearned_exclusion_1 = ""
	  unearned_exclusion_2 = ""
	  unearned_exclusion_3 = ""
'------------------------------

If (earned_income_request = 1  and unearned_income_request = 0) or (earned_income_request = 0  and unearned_income_request = 1) then
  income_dialog_height = 97
  income_button_height = 80
  
  if earned_income_request = 1 and unearned_income_request = 0 then
    earned_heading = 4
    earned_col_one_heading = 20
    earned_col_two_heading = 20
    earned_col_three_heading = 20
    earned_number_1 = 32
    earned_number_2 = 47
    earned_number_3 = 62
    earned_amnt_1 = 30
    earned_amnt_2 = 45
    earned_amnt_3 = 60
    earned_income_type_1 = 30
    earned_income_type_2 = 45
    earned_income_type_3 = 60
    earned_exclude_select_1 = 30
    earned_exclude_select_2 = 45
    earned_exclude_select_3 = 60
  ElseIf earned_income_request = 0  and unearned_income_request = 1 then
    unearned_heading = 4
    unearned_col_one_heading = 20
    unearned_col_two_heading = 20
    unearned_col_three_heading = 20
    unearned_number_1 = 32
    unearned_number_2 = 47
    unearned_number_3 = 62
    unearned_amnt_1 = 30
    unearned_amnt_2 = 45
    unearned_amnt_3 = 60
    unearned_income_type_1 = 30
    unearned_income_type_2 = 45
    unearned_income_type_3 = 60
    unearned_exclude_select_1 = 30
    unearned_exclude_select_2 = 45
    unearned_exclude_select_3 = 60
 End If 
  
ElseIf earned_income_request = 1 and unearned_income_request = 1 then
  'Sets Dynamic Window Parameters
  income_dialog_height = 177
  income_button_height = 160
  earned_heading = 4
  earned_col_one_heading = 20
  earned_col_two_heading = 20
  earned_col_three_heading = 20
  earned_number_1 = 32
  earned_number_2 = 47
  earned_number_3 = 62
  earned_amnt_1 = 30
  earned_amnt_2 = 45
  earned_amnt_3 = 60
  earned_income_type_1 = 30
  earned_income_type_2 = 45
  earned_income_type_3 = 60
  earned_exclude_select_1 = 30
  earned_exclude_select_2 = 45
  earned_exclude_select_3 = 60
  unearned_heading = 84
  unearned_col_one_heading = 100
  unearned_col_two_heading = 100
  unearned_col_three_heading = 100
  unearned_number_1 = 112
  unearned_number_2 = 127
  unearned_number_3 = 142
  unearned_amnt_1 = 110
  unearned_amnt_2 = 125
  unearned_amnt_3 = 140
  unearned_income_type_1 = 110
  unearned_income_type_2 = 125
  unearned_income_type_3 = 140
  unearned_exclude_select_1 = 110
  unearned_exclude_select_2 = 125
  unearned_exclude_select_3 = 140
  

End IF

BeginDialog addition_income_info_dialog, 0, 0, 220, income_dialog_height, "Income information"
  If earned_income_request = 1 and unearned_income_request = 0 then
    'Earned Income  
  Text 3, earned_heading, 209, 8, "Earned Income for member  "&client_name&"  in  "&grab_budget_month&"/"&grab_budget_year
  Text 17, earned_col_one_heading, 59, 8, "Amount per month"
  EditBox 15, earned_amnt_1, 80, 12, earned_value_1
  EditBox 15, earned_amnt_2, 80, 12, earned_value_2
  EditBox 15, earned_amnt_3, 80, 12, earned_value_3
  Text 10, earned_number_1, 5, 8, "1"
  Text 10, earned_number_2, 5, 8, "2"
  Text 10, earned_number_3, 5, 8, "3"
  DropListBox 100, earned_income_type_1, 80, 12, "Wages (Incl Tips)"+chr(9)+"WIA"+chr(9)+"EITC"+chr(9)+"Experiance Works"+chr(9)+"Federal Work Study"+chr(9)+"State Work Study"+chr(9)+"Other (JOBS)"+chr(9)+"Infrequent < %30 n/Recur"+chr(9)+"Infrequent <= $10 MSA Exclusion"+chr(9)+"Contract Income"+chr(9)+"Farm Income"+chr(9)+"Real Estate"+chr(9)+"Home Product Sales"+chr(9)+"Other Sales"+chr(9)+"Personal Services"+chr(9)+"Paper Route"+chr(9)+"In Home Daycare"+chr(9)+"Rental Income"+chr(9)+"Other (BUSI)"+chr(9)+"Roomer/Boarder Income"+chr(9)+"Boarder Income"+chr(9)+"Roomer Income", earned_type_1
  DropListBox 100, earned_income_type_2, 80, 12, "Wages (Incl Tips)"+chr(9)+"WIA"+chr(9)+"EITC"+chr(9)+"Experiance Works"+chr(9)+"Federal Work Study"+chr(9)+"State Work Study"+chr(9)+"Other (JOBS)"+chr(9)+"Infrequent < %30 n/Recur"+chr(9)+"Infrequent <= $10 MSA Exclusion"+chr(9)+"Contract Income"+chr(9)+"Farm Income"+chr(9)+"Real Estate"+chr(9)+"Home Product Sales"+chr(9)+"Other Sales"+chr(9)+"Personal Services"+chr(9)+"Paper Route"+chr(9)+"In Home Daycare"+chr(9)+"Rental Income"+chr(9)+"Other (BUSI)"+chr(9)+"Roomer/Boarder Income"+chr(9)+"Boarder Income"+chr(9)+"Roomer Income", earned_type_2
  DropListBox 100, earned_income_type_3, 80, 12, "Wages (Incl Tips)"+chr(9)+"WIA"+chr(9)+"EITC"+chr(9)+"Experiance Works"+chr(9)+"Federal Work Study"+chr(9)+"State Work Study"+chr(9)+"Other (JOBS)"+chr(9)+"Infrequent < %30 n/Recur"+chr(9)+"Infrequent <= $10 MSA Exclusion"+chr(9)+"Contract Income"+chr(9)+"Farm Income"+chr(9)+"Real Estate"+chr(9)+"Home Product Sales"+chr(9)+"Other Sales"+chr(9)+"Personal Services"+chr(9)+"Paper Route"+chr(9)+"In Home Daycare"+chr(9)+"Rental Income"+chr(9)+"Other (BUSI)"+chr(9)+"Roomer/Boarder Income"+chr(9)+"Boarder Income"+chr(9)+"Roomer Income", earned_type_3
  Text 102, earned_col_two_heading, 50, 8, "Income Type"
  Text 185, earned_col_three_heading, 50, 8, "Excluded"
  DropListBox 185, earned_exclude_select_1, 30, 12, "No"+chr(9)+"Yes", earned_exclusion_1
  DropListBox 185, earned_exclude_select_2, 30, 12, "No"+chr(9)+"Yes", earned_exclusion_2
  DropListBox 185, earned_exclude_select_3, 30, 12, "No"+chr(9)+"Yes", earned_exclusion_3
  ButtonGroup ButtonPressed
    OkButton 160, income_button_height, 20, 12
    CancelButton 185, income_button_height, 30, 12
  ElseIf earned_income_request = 0 and unearned_income_request = 1 then
    'Unearned Income 
  Text 3, unearned_heading, 209, 8, "Unearned Income for member  "&client_name&"  in  "&grab_budget_month&"/"&grab_budget_year
  Text 17, unearned_col_one_heading, 59, 8, "Amount per month"
  EditBox 15, unearned_amnt_1, 80, 12, unearned_value_1
  EditBox 15, unearned_amnt_2, 80, 12, unearned_value_2
  EditBox 15, unearned_amnt_3, 80, 12, unearned_value_3
  Text 10, unearned_number_1, 5, 8, "1"
  Text 10, unearned_number_2, 5, 8, "2"
  Text 10, unearned_number_3, 5, 8, "3"
  DropListBox 100, unearned_income_type_1, 80, 12, "RSDI, Disa"+chr(9)+"RSDI, Non-Disa"+chr(9)+"SSI"+chr(9)+"Direct Spousal Support"+chr(9)+"Disbursed Child Support"+chr(9)+"Non-MA PA"+chr(9)+"Disbursed Spousal Support"+chr(9)+"Direct Child Support"+chr(9)+"VA Disability"+chr(9)+"VA Pension"+chr(9)+"VA Other"+chr(9)+"Unemployment Insurance"+chr(9)+"Worker's Comp"+chr(9)+"Railroad Retirement"+chr(9)+"Other Retirement"+chr(9)+"Military Allotment"+chr(9)+"FC Child Requesting FS"+chr(9)+"FC Child Not Requesting FS"+chr(9)+"FC Adult Requesting FS"+chr(9)+"FC Adult Not Requesting FS"+chr(9)+"Dividends"+chr(9)+"Interests"+chr(9)+"Cnt Gifts or Prized"+chr(9)+"Strike Benefit"+chr(9)+"Contract for Deed"+chr(9)+"Illegal Income"+chr(9)+"Other Countable"+chr(9)+"Infrequent < 30 Not Counted"+chr(9)+"Other, FS Only"+chr(9)+"Infreq <= $20 MSA Exclustion"+chr(9)+"Rental Income"+chr(9)+"Student, Non Title-IV Aid"+chr(9)+"Student, Title-IV Aid"+chr(9)+"VA Aid & Attendance"+chr(9)+"Disbursed CS Arrears"+chr(9)+"Disbursed Spsl Sup Arrears"+chr(9)+"MSA"+chr(9)+"GA"+chr(9)+"RCA"+chr(9)+"MFIP"+chr(9)+"Lump Sum"+chr(9)+"Disbursed Excess CS", unearned_type_1
  DropListBox 100, unearned_income_type_2, 80, 12, "RSDI, Disa"+chr(9)+"RSDI, Non-Disa"+chr(9)+"SSI"+chr(9)+"Direct Spousal Support"+chr(9)+"Disbursed Child Support"+chr(9)+"Non-MA PA"+chr(9)+"Disbursed Spousal Support"+chr(9)+"Direct Child Support"+chr(9)+"VA Disability"+chr(9)+"VA Pension"+chr(9)+"VA Other"+chr(9)+"Unemployment Insurance"+chr(9)+"Worker's Comp"+chr(9)+"Railroad Retirement"+chr(9)+"Other Retirement"+chr(9)+"Military Allotment"+chr(9)+"FC Child Requesting FS"+chr(9)+"FC Child Not Requesting FS"+chr(9)+"FC Adult Requesting FS"+chr(9)+"FC Adult Not Requesting FS"+chr(9)+"Dividends"+chr(9)+"Interests"+chr(9)+"Cnt Gifts or Prized"+chr(9)+"Strike Benefit"+chr(9)+"Contract for Deed"+chr(9)+"Illegal Income"+chr(9)+"Other Countable"+chr(9)+"Infrequent < 30 Not Counted"+chr(9)+"Other, FS Only"+chr(9)+"Infreq <= $20 MSA Exclustion"+chr(9)+"Rental Income"+chr(9)+"Student, Non Title-IV Aid"+chr(9)+"Student, Title-IV Aid"+chr(9)+"VA Aid & Attendance"+chr(9)+"Disbursed CS Arrears"+chr(9)+"Disbursed Spsl Sup Arrears"+chr(9)+"MSA"+chr(9)+"GA"+chr(9)+"RCA"+chr(9)+"MFIP"+chr(9)+"Lump Sum"+chr(9)+"Disbursed Excess CS", unearned_type_2
  DropListBox 100, unearned_income_type_3, 80, 12, "RSDI, Disa"+chr(9)+"RSDI, Non-Disa"+chr(9)+"SSI"+chr(9)+"Direct Spousal Support"+chr(9)+"Disbursed Child Support"+chr(9)+"Non-MA PA"+chr(9)+"Disbursed Spousal Support"+chr(9)+"Direct Child Support"+chr(9)+"VA Disability"+chr(9)+"VA Pension"+chr(9)+"VA Other"+chr(9)+"Unemployment Insurance"+chr(9)+"Worker's Comp"+chr(9)+"Railroad Retirement"+chr(9)+"Other Retirement"+chr(9)+"Military Allotment"+chr(9)+"FC Child Requesting FS"+chr(9)+"FC Child Not Requesting FS"+chr(9)+"FC Adult Requesting FS"+chr(9)+"FC Adult Not Requesting FS"+chr(9)+"Dividends"+chr(9)+"Interests"+chr(9)+"Cnt Gifts or Prized"+chr(9)+"Strike Benefit"+chr(9)+"Contract for Deed"+chr(9)+"Illegal Income"+chr(9)+"Other Countable"+chr(9)+"Infrequent < 30 Not Counted"+chr(9)+"Other, FS Only"+chr(9)+"Infreq <= $20 MSA Exclustion"+chr(9)+"Rental Income"+chr(9)+"Student, Non Title-IV Aid"+chr(9)+"Student, Title-IV Aid"+chr(9)+"VA Aid & Attendance"+chr(9)+"Disbursed CS Arrears"+chr(9)+"Disbursed Spsl Sup Arrears"+chr(9)+"MSA"+chr(9)+"GA"+chr(9)+"RCA"+chr(9)+"MFIP"+chr(9)+"Lump Sum"+chr(9)+"Disbursed Excess CS", unearned_type_3
  Text 102, unearned_col_two_heading, 50, 8, "Income Type"
  Text 185, unearned_col_three_heading, 50, 8, "Excluded"
  DropListBox 185, unearned_exclude_select_1, 30, 12, "No"+chr(9)+"Yes", unearned_exclusion_1
  DropListBox 185, unearned_exclude_select_2, 30, 12, "No"+chr(9)+"Yes", unearned_exclusion_2
  DropListBox 185, unearned_exclude_select_3, 30, 12, "No"+chr(9)+"Yes", unearned_exclusion_3
  ButtonGroup ButtonPressed
    OkButton 160, income_button_height, 20, 12
    CancelButton 185, income_button_height, 30, 12
  ElseIf earned_income_request = 1 and unearned_income_request = 1 then
    'Earned Income
  Text 3, earned_heading, 209, 8, "Earned Income for "&client_name
  Text 17, earned_col_one_heading, 59, 8, "Amount per month"
  EditBox 15, earned_amnt_1, 80, 12, earned_value_1
  EditBox 15, earned_amnt_2, 80, 12, earned_value_2
  EditBox 15, earned_amnt_3, 80, 12, earned_value_3
  Text 10, earned_number_1, 5, 8, "1"
  Text 10, earned_number_2, 5, 8, "2"
  Text 10, earned_number_3, 5, 8, "3"
  DropListBox 100, earned_income_type_1, 80, 12, "Wages (Incl Tips)"+chr(9)+"WIA"+chr(9)+"EITC"+chr(9)+"Experiance Works"+chr(9)+"Federal Work Study"+chr(9)+"State Work Study"+chr(9)+"Other (JOBS)"+chr(9)+"Infrequent < %30 n/Recur"+chr(9)+"Infrequent <= $10 MSA Exclusion"+chr(9)+"Contract Income"+chr(9)+"Farm Income"+chr(9)+"Real Estate"+chr(9)+"Home Product Sales"+chr(9)+"Other Sales"+chr(9)+"Personal Services"+chr(9)+"Paper Route"+chr(9)+"In Home Daycare"+chr(9)+"Rental Income"+chr(9)+"Other (BUSI)"+chr(9)+"Roomer/Boarder Income"+chr(9)+"Boarder Income"+chr(9)+"Roomer Income", earned_type_1
  DropListBox 100, earned_income_type_2, 80, 12, "Wages (Incl Tips)"+chr(9)+"WIA"+chr(9)+"EITC"+chr(9)+"Experiance Works"+chr(9)+"Federal Work Study"+chr(9)+"State Work Study"+chr(9)+"Other (JOBS)"+chr(9)+"Infrequent < %30 n/Recur"+chr(9)+"Infrequent <= $10 MSA Exclusion"+chr(9)+"Contract Income"+chr(9)+"Farm Income"+chr(9)+"Real Estate"+chr(9)+"Home Product Sales"+chr(9)+"Other Sales"+chr(9)+"Personal Services"+chr(9)+"Paper Route"+chr(9)+"In Home Daycare"+chr(9)+"Rental Income"+chr(9)+"Other (BUSI)"+chr(9)+"Roomer/Boarder Income"+chr(9)+"Boarder Income"+chr(9)+"Roomer Income", earned_type_2
  DropListBox 100, earned_income_type_3, 80, 12, "Wages (Incl Tips)"+chr(9)+"WIA"+chr(9)+"EITC"+chr(9)+"Experiance Works"+chr(9)+"Federal Work Study"+chr(9)+"State Work Study"+chr(9)+"Other (JOBS)"+chr(9)+"Infrequent < %30 n/Recur"+chr(9)+"Infrequent <= $10 MSA Exclusion"+chr(9)+"Contract Income"+chr(9)+"Farm Income"+chr(9)+"Real Estate"+chr(9)+"Home Product Sales"+chr(9)+"Other Sales"+chr(9)+"Personal Services"+chr(9)+"Paper Route"+chr(9)+"In Home Daycare"+chr(9)+"Rental Income"+chr(9)+"Other (BUSI)"+chr(9)+"Roomer/Boarder Income"+chr(9)+"Boarder Income"+chr(9)+"Roomer Income", earned_type_3
  Text 102, earned_col_two_heading, 50, 8, "Income Type"
  Text 185, earned_col_three_heading, 50, 8, "Excluded"
  DropListBox 185, earned_exclude_select_1, 30, 12, "No"+chr(9)+"Yes", earned_exclusion_1
  DropListBox 185, earned_exclude_select_2, 30, 12, "No"+chr(9)+"Yes", earned_exclusion_2
  DropListBox 185, earned_exclude_select_3, 30, 12, "No"+chr(9)+"Yes", earned_exclusion_3
    'Unearned Income 
  Text 3, unearned_heading, 209, 8, "Unearned Income for "&client_name
  Text 17, unearned_col_one_heading, 59, 8, "Amount per month"
  EditBox 15, unearned_amnt_1, 80, 12, unearned_value_1
  EditBox 15, unearned_amnt_2, 80, 12, unearned_value_2
  EditBox 15, unearned_amnt_3, 80, 12, unearned_value_3
  Text 10, unearned_number_1, 5, 8, "1"
  Text 10, unearned_number_2, 5, 8, "2"
  Text 10, unearned_number_3, 5, 8, "3"
  DropListBox 100, unearned_income_type_1, 80, 12, "RSDI, Disa"+chr(9)+"RSDI, Non-Disa"+chr(9)+"SSI"+chr(9)+"Direct Spousal Support"+chr(9)+"Disbursed Child Support"+chr(9)+"Non-MA PA"+chr(9)+"Disbursed Spousal Support"+chr(9)+"Direct Child Support"+chr(9)+"VA Disability"+chr(9)+"VA Pension"+chr(9)+"VA Other"+chr(9)+"Unemployment Insurance"+chr(9)+"Worker's Comp"+chr(9)+"Railroad Retirement"+chr(9)+"Other Retirement"+chr(9)+"Military Allotment"+chr(9)+"FC Child Requesting FS"+chr(9)+"FC Child Not Requesting FS"+chr(9)+"FC Adult Requesting FS"+chr(9)+"FC Adult Not Requesting FS"+chr(9)+"Dividends"+chr(9)+"Interests"+chr(9)+"Cnt Gifts or Prized"+chr(9)+"Strike Benefit"+chr(9)+"Contract for Deed"+chr(9)+"Illegal Income"+chr(9)+"Other Countable"+chr(9)+"Infrequent < 30 Not Counted"+chr(9)+"Other, FS Only"+chr(9)+"Infreq <= $20 MSA Exclustion"+chr(9)+"Rental Income"+chr(9)+"Student, Non Title-IV Aid"+chr(9)+"Student, Title-IV Aid"+chr(9)+"VA Aid & Attendance"+chr(9)+"Disbursed CS Arrears"+chr(9)+"Disbursed Spsl Sup Arrears"+chr(9)+"MSA"+chr(9)+"GA"+chr(9)+"RCA"+chr(9)+"MFIP"+chr(9)+"Lump Sum"+chr(9)+"Disbursed Excess CS", unearned_type_1
  DropListBox 100, unearned_income_type_2, 80, 12, "RSDI, Disa"+chr(9)+"RSDI, Non-Disa"+chr(9)+"SSI"+chr(9)+"Direct Spousal Support"+chr(9)+"Disbursed Child Support"+chr(9)+"Non-MA PA"+chr(9)+"Disbursed Spousal Support"+chr(9)+"Direct Child Support"+chr(9)+"VA Disability"+chr(9)+"VA Pension"+chr(9)+"VA Other"+chr(9)+"Unemployment Insurance"+chr(9)+"Worker's Comp"+chr(9)+"Railroad Retirement"+chr(9)+"Other Retirement"+chr(9)+"Military Allotment"+chr(9)+"FC Child Requesting FS"+chr(9)+"FC Child Not Requesting FS"+chr(9)+"FC Adult Requesting FS"+chr(9)+"FC Adult Not Requesting FS"+chr(9)+"Dividends"+chr(9)+"Interests"+chr(9)+"Cnt Gifts or Prized"+chr(9)+"Strike Benefit"+chr(9)+"Contract for Deed"+chr(9)+"Illegal Income"+chr(9)+"Other Countable"+chr(9)+"Infrequent < 30 Not Counted"+chr(9)+"Other, FS Only"+chr(9)+"Infreq <= $20 MSA Exclustion"+chr(9)+"Rental Income"+chr(9)+"Student, Non Title-IV Aid"+chr(9)+"Student, Title-IV Aid"+chr(9)+"VA Aid & Attendance"+chr(9)+"Disbursed CS Arrears"+chr(9)+"Disbursed Spsl Sup Arrears"+chr(9)+"MSA"+chr(9)+"GA"+chr(9)+"RCA"+chr(9)+"MFIP"+chr(9)+"Lump Sum"+chr(9)+"Disbursed Excess CS", unearned_type_2
  DropListBox 100, unearned_income_type_3, 80, 12, "RSDI, Disa"+chr(9)+"RSDI, Non-Disa"+chr(9)+"SSI"+chr(9)+"Direct Spousal Support"+chr(9)+"Disbursed Child Support"+chr(9)+"Non-MA PA"+chr(9)+"Disbursed Spousal Support"+chr(9)+"Direct Child Support"+chr(9)+"VA Disability"+chr(9)+"VA Pension"+chr(9)+"VA Other"+chr(9)+"Unemployment Insurance"+chr(9)+"Worker's Comp"+chr(9)+"Railroad Retirement"+chr(9)+"Other Retirement"+chr(9)+"Military Allotment"+chr(9)+"FC Child Requesting FS"+chr(9)+"FC Child Not Requesting FS"+chr(9)+"FC Adult Requesting FS"+chr(9)+"FC Adult Not Requesting FS"+chr(9)+"Dividends"+chr(9)+"Interests"+chr(9)+"Cnt Gifts or Prized"+chr(9)+"Strike Benefit"+chr(9)+"Contract for Deed"+chr(9)+"Illegal Income"+chr(9)+"Other Countable"+chr(9)+"Infrequent < 30 Not Counted"+chr(9)+"Other, FS Only"+chr(9)+"Infreq <= $20 MSA Exclustion"+chr(9)+"Rental Income"+chr(9)+"Student, Non Title-IV Aid"+chr(9)+"Student, Title-IV Aid"+chr(9)+"VA Aid & Attendance"+chr(9)+"Disbursed CS Arrears"+chr(9)+"Disbursed Spsl Sup Arrears"+chr(9)+"MSA"+chr(9)+"GA"+chr(9)+"RCA"+chr(9)+"MFIP"+chr(9)+"Lump Sum"+chr(9)+"Disbursed Excess CS", unearned_type_3
  Text 102, unearned_col_two_heading, 50, 8, "Income Type"
  Text 185, unearned_col_three_heading, 50, 8, "Excluded"
  DropListBox 185, unearned_exclude_select_1, 30, 12, "No"+chr(9)+"Yes", unearned_exclusion_1
  DropListBox 185, unearned_exclude_select_2, 30, 12, "No"+chr(9)+"Yes", unearned_exclusion_2
  DropListBox 185, unearned_exclude_select_3, 30, 12, "No"+chr(9)+"Yes", unearned_exclusion_3
  ButtonGroup ButtonPressed
    OkButton 160, income_button_height, 20, 12
    CancelButton 185, income_button_height, 30, 12 
  End If 
EndDialog

  Dialog addition_income_info_dialog
    If buttonpressed = 0 then stopscript

  earned_type_1 = trim(earned_type_1)
  earned_type_2 = trim(earned_type_2)
  earned_type_3 = trim(earned_type_3)
  unearned_type_1 = trim(unearned_type_1)
  unearned_type_2 = trim(unearned_type_2)
  unearned_type_3 = trim(unearned_type_3)
  earned_value_1 = trim(earned_value_1)
  earned_value_2 = trim(earned_value_2)
  earned_value_3 = trim(earned_value_3)
  unearned_value_1 = trim(unearned_value_1)
  unearned_value_2 = trim(unearned_value_2)
  unearned_value_3 = trim(unearned_value_3)
  earned_exclusion_1 = trim(earned_exclusion_1)
  earned_exclusion_2 = trim(earned_exclusion_2)
  earned_exclusion_3 = trim(earned_exclusion_3)
  unearned_exclusion_1 = trim(unearned_exclusion_1)
  unearned_exclusion_2 = trim(unearned_exclusion_2)
  unearned_exclusion_3 = trim(unearned_exclusion_3)

End Function

Function income_type_converter(title_of_income_type)
  converted_income_type = ""
  title_of_income_type = trim(title_of_income_type)
  If title_of_income_type = "WIA" then title_of_converted_income_type = "01"
  If title_of_income_type = "Wages (Incl Tips)" then title_of_converted_income_type = "02"
  If title_of_income_type = "EITC" then title_of_converted_income_type = "03"
  If title_of_income_type = "Experiance Works" then title_of_converted_income_type = "04"
  If title_of_income_type = "Federal Work Study" then title_of_converted_income_type = "05"
  If title_of_income_type = "State Work Study" then title_of_converted_income_type = "06"
  If title_of_income_type = "Other (JOBS)" then title_of_converted_income_type = "07"
  If title_of_income_type = "Infrequent < %30 n/Recur" then title_of_converted_income_type = "08"
  If title_of_income_type = "Infrequent <= $10 MSA Exclusion" then title_of_converted_income_type = "09"
  If title_of_income_type = "Contract Income" then title_of_converted_income_type = "10"
  If title_of_income_type = "Farm Income" then title_of_converted_income_type = "11"
  If title_of_income_type = "Real Estate" then title_of_converted_income_type = "14"
  If title_of_income_type = "Home Product Sales" then title_of_converted_income_type = "15"
  If title_of_income_type = "Other Sales" then title_of_converted_income_type = "16"
  If title_of_income_type = "Personal Services" then title_of_converted_income_type = "17"
  If title_of_income_type = "Paper Route" then title_of_converted_income_type = "18"
  If title_of_income_type = "In Home Daycare" then title_of_converted_income_type = "19"
  If title_of_income_type = "Rental Income" then title_of_converted_income_type = "20"
  If title_of_income_type = "Other (BUSI)" then title_of_converted_income_type = "21"
  If title_of_income_type = "Roomer/Boarder Income" then title_of_converted_income_type = "22"
  If title_of_income_type = "Boarder Income" then title_of_converted_income_type = "23"
  If title_of_income_type = "Roomer Income" then title_of_converted_income_type = "24"
  If title_of_income_type = "RSDI, Disa" then title_of_converted_income_type = "01"
  If title_of_income_type = "RSDI, Non-Disa" then title_of_converted_income_type = "02"
  If title_of_income_type = "SSI" then title_of_converted_income_type = "03"
  If title_of_income_type = "Direct Spousal Support" then title_of_income_converted_type = "04"
  If title_of_income_type = "Disbursed Child Support" then title_of_income_converted_type = "05"
  If title_of_income_type = "Non-MN PA" then title_of_converted_income_type = "06"
  If title_of_income_type = "Disbursed Spousal Support" then title_of_income_converted_type = "07"
  If title_of_income_type = "Direct Child Support" then title_of_converted_income_type = "08"
  If title_of_income_type = "VA Disability" then title_of_converted_income_type = "09"
  If title_of_income_type = "VA Pension" then title_of_converted_income_type = "10"
  If title_of_income_type = "VA Other" then title_of_converted_income_type = "11"
  If title_of_income_type = "Unemployment Insurance" then title_of_converted_income_type = "12"
  If title_of_income_type = "Worker's Comp" then title_of_converted_income_type = "13"
  If title_of_income_type = "Railroad Retirement" then title_of_converted_income_type = "14"
  If title_of_income_type = "Other Retirement" then title_of_converted_income_type = "15"
  If title_of_income_type = "Military Allotment" then title_of_converted_income_type = "16"
  If title_of_income_type = "FC Child Requesting FS" then title_of_converted_income_type = "17"
  If title_of_income_type = "FC Child Not Requesting FS" then title_of_converted_income_type = "18"
  If title_of_income_type = "FC Adult Requesting FS" then title_of_converted_income_type = "19"
  If title_of_income_type = "FC Adult Not Requesting FS" then title_of_converted_income_type = "20"
  If title_of_income_type = "Dividends" then title_of_converted_income_type = "21"
  If title_of_income_type = "Interests" then title_of_converted_income_type = "22"
  If title_of_income_type = "Cnt Gifts or Prized" then title_of_converted_income_type = "23"
  If title_of_income_type = "Strike Benefit" then title_of_converted_income_type = "24"
  If title_of_income_type = "Contract for Deed" then title_of_converted_income_type = "25"
  If title_of_income_type = "Illegal Income" then title_of_converted_income_type = "26"
  If title_of_income_type = "Other Countable" then title_of_converted_income_type = "27"
  If title_of_income_type = "Infrequent < 30 Not Counted" then title_of_converted_income_type = "28"
  If title_of_income_type = "Other, FS Only" then title_of_converted_income_type = "29"
  If title_of_income_type = "Infreq <= $20 MSA Exclustion" then title_of_converted_income_type = "30"
  If title_of_income_type = "Rental Income" then title_of_converted_income_type = "31"
  If title_of_income_type = "Student, Non Title-IV Aid" then title_of_converted_income_type = "32"
  If title_of_income_type = "Student, Title-IV Aid" then title_of_converted_income_type = "33"
  If title_of_income_type = "VA Aid & Attendance" then title_of_converted_income_type = "34"
  If title_of_income_type = "Disbursed CS Arrears" then title_of_converted_income_type = "35"
  If title_of_income_type = "Disbursed Spsl Sup Arrears" then title_of_converted_income_type = "36"
  If title_of_income_type = "MSA" then title_of_converted_income_type = "38"
  If title_of_income_type = "GA" then title_of_converted_income_type = "39"
  If title_of_income_type = "RCA" then title_of_converted_income_type = "40"
  If title_of_income_type = "MFIP" then title_of_converted_income_type = "41"
  If title_of_income_type = "Lump Sum" then title_of_converted_income_type = "42"
  If title_of_income_type = "Disbursed Excess CS" then title_of_converted_income_type = "43"
  converted_income_type = title_of_converted_income_type
End Function

function enter_income_screen_information(searchable_client_name,client_name_is,current_client_selected,earned_type_1,earned_type_2,earned_type_3,unearned_type_1,unearned_type_2,unearned_type_3,earned_value_1,earned_value_2,earned_value_3,unearned_value_1,unearned_value_2,unearned_value_3,earned_exclusion_1,earned_exclusion_2,earned_exclusion_3,unearned_exclusion_1,unearned_exclusion_2,unearned_exclusion_3)
  converted_income_type = ""
  earned_type_1 = trim(earned_type_1)
  earned_type_2 = trim(earned_type_2)
  earned_type_3 = trim(earned_type_3)
  unearned_type_1 = trim(unearned_type_1)
  unearned_type_2 = trim(unearned_type_2)
  unearned_type_3 = trim(unearned_type_3)
  earned_value_1 = trim(earned_value_1)
  earned_value_2 = trim(earned_value_2)
  earned_value_3 = trim(earned_value_3)
  unearned_value_1 = trim(unearned_value_1)
  unearned_value_2 = trim(unearned_value_2)
  unearned_value_3 = trim(unearned_value_3)
  earned_exclusion_1 = trim(earned_exclusion_1)
  earned_exclusion_2 = trim(earned_exclusion_2)
  earned_exclusion_3 = trim(earned_exclusion_3)
  unearned_exclusion_1 = trim(unearned_exclusion_1)
  unearned_exclusion_2 = trim(unearned_exclusion_2)
  unearned_exclusion_3 = trim(unearned_exclusion_3)  
  If (earned_value_1 <> "" or unearned_value_1 <> "") and client_name_is = current_client_selected then
	If earned_value_1 <> "" then
	  EMReadScreen elig_abud_check, 4, 3, 47
	    elig_abud_check = trim(elig_abud_check)
	  EMReadScreen elig_bbud_check, 4, 3, 47
      elig_bbud_check = trim(elig_bbud_check)
    EMReadScreen elig_cbud_check, 4, 3, 54
	    elig_cbud_check = trim(elig_cbud_check)
	  If elig_abud_check = "ABUD" then 
	    EMWriteScreen "X", 14, 3
    ElseIf elig_bbud_check = "BBUD" then
      EMWriteScreen "X", 8, 43
	  ElseIf elig_cbud_check = "CBUD" then
      EMWriteScreen "X", 8, 43
	  End If

	  transmit
	  call income_type_converter(earned_type_1)
	  EMWriteScreen "__", 8, 8
	  EMWriteScreen "___________", 8, 43
	  EMWriteScreen "_", 8, 59
	  EMWriteScreen converted_income_type, 8, 8
	  EMWriteScreen earned_value_1, 8, 43
	  EMWriteScreen earned_exclusion_1, 8, 59
	  transmit
	  PF3
	End If
	If earned_value_2 <> "" then
	  EMReadScreen elig_abud_check, 4, 3, 47
	    elig_abud_check = trim(elig_abud_check)
	  EMReadScreen elig_bbud_check, 4, 3, 47
      elig_bbud_check = trim(elig_bbud_check)
    EMReadScreen elig_cbud_check, 4, 3, 54
	    elig_cbud_check = trim(elig_cbud_check)
	  If elig_abud_check = "ABUD" then 
	    EMWriteScreen "X", 14, 3
    ElseIf elig_bbud_check = "BBUD" then
      EMWriteScreen "X", 8, 43
	  ElseIf elig_cbud_check = "CBUD" then
      EMWriteScreen "X", 8, 43
	  End If

	  transmit
	  call income_type_converter(earned_type_2)
	  EMWriteScreen "__", 9, 8
	  EMWriteScreen "___________", 9, 43
	  EMWriteScreen "_", 9, 59
	  EMWriteScreen converted_income_type, 9, 8
	  EMWriteScreen earned_value_2, 9, 43
	  EMWriteScreen earned_exclusion_2, 9, 59
	  transmit
	  PF3
	End If
	If earned_value_3 <> "" then
	  EMReadScreen elig_abud_check, 4, 3, 47
	    elig_abud_check = trim(elig_abud_check)
	  EMReadScreen elig_bbud_check, 4, 3, 47
      elig_bbud_check = trim(elig_bbud_check)
    EMReadScreen elig_cbud_check, 4, 3, 54
	    elig_cbud_check = trim(elig_cbud_check)
	  If elig_abud_check = "ABUD" then 
	    EMWriteScreen "X", 14, 3
    ElseIf elig_bbud_check = "BBUD" then
      EMWriteScreen "X", 8, 43
	  ElseIf elig_cbud_check = "CBUD" then
      EMWriteScreen "X", 8, 43
	  End If

	  transmit
	  call income_type_converter(earned_type_3)
	  EMWriteScreen "__", 10, 8
	  EMWriteScreen "___________", 10, 43
	  EMWriteScreen "_", 10, 59
	  EMWriteScreen converted_income_type, 10, 8
	  EMWriteScreen earned_value_3, 10, 43
	  EMWriteScreen earned_exclusion_3, 10, 59
	  transmit
	  PF3
	End If
	If unearned_value_1 <> "" then
  	EMReadScreen elig_abud_check, 4, 3, 47
	    elig_abud_check = trim(elig_abud_check)
    EMReadScreen elig_bbud_check, 4, 3, 47  
      elig_bbud_check = trim(elig_bbud_check)
	  EMReadScreen elig_cbud_check, 4, 3, 54
	    elig_cbud_check = trim(elig_cbud_check)
	  If elig_abud_check = "ABUD" then 
	    EMWriteScreen "X", 9, 3
    ElseIf elig_bbud_check = "BBUD" then
      EMWRITEScreen "X", 8, 3
	  ElseIf elig_cbud_check = "CBUD" then
		EMWriteScreen "X", 8, 3
	  End If

      transmit
	  call income_type_converter(unearned_type_1)
	  EMWriteScreen "__", 8, 8
	  EMWriteScreen "___________", 8, 43
	  EMWriteScreen "_", 8, 58
	  EMWriteScreen converted_income_type, 8, 8
      EMWriteScreen unearned_value_1, 8, 43
      EMWriteScreen unearned_exclusion_1, 8, 58
      transmit
      PF3	  
	End If
	If unearned_value_2 <> "" then
  	EMReadScreen elig_abud_check, 4, 3, 47
	    elig_abud_check = trim(elig_abud_check)
    EMReadScreen elig_bbud_check, 4, 3, 47  
      elig_bbud_check = trim(elig_bbud_check)
	  EMReadScreen elig_cbud_check, 4, 3, 54
	    elig_cbud_check = trim(elig_cbud_check)
	  If elig_abud_check = "ABUD" then 
	    EMWriteScreen "X", 9, 3
    ElseIf elig_bbud_check = "BBUD" then
      EMWRITEScreen "X", 8, 3
	  ElseIf elig_cbud_check = "CBUD" then
		EMWriteScreen "X", 8, 3
	  End If
	  
      transmit
	  call income_type_converter(unearned_type_2)
	  EMWriteScreen "__", 9, 8
	  EMWriteScreen "___________", 9, 43
	  EMWriteScreen "_", 9, 58
	  EMWriteScreen converted_income_type, 9, 8
      EMWriteScreen unearned_value_2, 9, 43
      EMWriteScreen unearned_exclusion_2, 9, 58
      transmit
      PF3	  
	End If
	If unearned_value_3 <> "" then
  	EMReadScreen elig_abud_check, 4, 3, 47
	    elig_abud_check = trim(elig_abud_check)
    EMReadScreen elig_bbud_check, 4, 3, 47  
      elig_bbud_check = trim(elig_bbud_check)
	  EMReadScreen elig_cbud_check, 4, 3, 54
	    elig_cbud_check = trim(elig_cbud_check)
	  If elig_abud_check = "ABUD" then 
	    EMWriteScreen "X", 9, 3
    ElseIf elig_bbud_check = "BBUD" then
      EMWRITEScreen "X", 8, 3
	  ElseIf elig_cbud_check = "CBUD" then
		EMWriteScreen "X", 8, 3
	  End If	  
      transmit
	  call income_type_converter(unearned_type_3)
	  EMWriteScreen "__", 10, 8
	  EMWriteScreen "___________", 10, 43
	  EMWriteScreen "_", 10, 58
	  EMWriteScreen converted_income_type, 10, 8
      EMWriteScreen unearned_value_3, 10, 43
      EMWriteScreen unearned_exclusion_3, 10, 58
      transmit
      PF3	  
	End If
  ElseIf (earned_value_1 <> "" or unearned_value_1 <> "") and client_name_is <> current_client_selected then
    If earned_value_1 <> "" then
	  EMReadScreen elig_abud_check, 4, 3, 47
	    elig_abud_check = trim(elig_abud_check)
	  EMReadScreen elig_bbud_check, 4, 3, 47
      elig_bbud_check = trim(elig_bbud_check)
    EMReadScreen elig_cbud_check, 4, 3, 54
	    elig_cbud_check = trim(elig_cbud_check)
	  If elig_abud_check = "ABUD" then 
	    EMWriteScreen "X", 13, 43
    ElseIf elig_bbud_check = "BBUD" then
      EMWriteScreen "X", 9, 43
	  ElseIf elig_cbud_check = "CBUD" then
      EMWriteScreen "X", 13, 43
	  End If
	  transmit
	  row = 1
	  col = 1
	  EMSearch "Ref Nbr: "&searchable_client_name, row, col
	  If row <> 0 then
	    EMWriteScreen "X", row + 6, col - 1
		transmit
	    call income_type_converter(earned_type_1)
		EMWriteScreen "__", 8, 80
	    EMWriteScreen "___________", 8, 43
	    EMWriteScreen "_", 8, 59
	    EMWriteScreen converted_income_type, 8, 8
	    EMWriteScreen earned_value_1, 8, 43
	    EMWriteScreen earned_exclusion_1, 8, 59
		transmit
		PF3
	  End If
	  PF3
	End If
	If earned_value_2 <> "" then
	  EMReadScreen elig_abud_check, 4, 3, 47
	    elig_abud_check = trim(elig_abud_check)
	  EMReadScreen elig_bbud_check, 4, 3, 47
      elig_bbud_check = trim(elig_bbud_check)
    EMReadScreen elig_cbud_check, 4, 3, 54
	    elig_cbud_check = trim(elig_cbud_check)
	  If elig_abud_check = "ABUD" then 
	    EMWriteScreen "X", 13, 43
    ElseIf elig_bbud_check = "BBUD" then
      EMWriteScreen "X", 9, 43
	  ElseIf elig_cbud_check = "CBUD" then
      EMWriteScreen "X", 13, 43
	  End If
	  transmit
	  row = 1
	  col = 1
	  EMSearch "Ref Nbr: "&searchable_client_name, row, col
	  If row <> 0 then
	    EMWriteScreen "X", row + 6, col - 1
		transmit
	    call income_type_converter(earned_type_2)
		EMWriteScreen "__", 9, 80
	    EMWriteScreen "___________", 9, 43
	    EMWriteScreen "_", 9, 59
	    EMWriteScreen converted_income_type, 9, 8
	    EMWriteScreen earned_value_2, 9, 43
	    EMWriteScreen earned_exclusion_2, 9, 59
		transmit
		PF3
	  End If
	  PF3
	End If
	If earned_value_3 <> "" then
	  EMReadScreen elig_abud_check, 4, 3, 47
	    elig_abud_check = trim(elig_abud_check)
	  EMReadScreen elig_bbud_check, 4, 3, 47
      elig_bbud_check = trim(elig_bbud_check)
    EMReadScreen elig_cbud_check, 4, 3, 54
	    elig_cbud_check = trim(elig_cbud_check)
	  If elig_abud_check = "ABUD" then 
	    EMWriteScreen "X", 13, 43
    ElseIf elig_bbud_check = "BBUD" then
      EMWriteScreen "X", 9, 43
	  ElseIf elig_cbud_check = "CBUD" then
      EMWriteScreen "X", 13, 43
	  End If
	  transmit
	  row = 1
	  col = 1
	  EMSearch "Ref Nbr: "&searchable_client_name, row, col
	  If row <> 0 then
	    EMWriteScreen "X", row + 6, col - 1
		transmit
	    call income_type_converter(earned_type_3)
		EMWriteScreen "__", 10, 80
	    EMWriteScreen "___________", 10, 43
	    EMWriteScreen "_", 10, 59
	    EMWriteScreen converted_income_type, 10, 8
	    EMWriteScreen earned_value_3, 10, 43
	    EMWriteScreen earned_exclusion_3, 10, 59
		transmit
		PF3
	  End If
	  PF3
	End If
	If unearned_value_1 <> "" then
  	EMReadScreen elig_abud_check, 4, 3, 47
	    elig_abud_check = trim(elig_abud_check)
    EMReadScreen elig_bbud_check, 4, 3, 47  
      elig_bbud_check = trim(elig_bbud_check)
	  EMReadScreen elig_cbud_check, 4, 3, 54
	    elig_cbud_check = trim(elig_cbud_check)
	  If elig_abud_check = "ABUD" then 
	    EMWriteScreen "X", 13, 43
    ElseIf elig_bbud_check = "BBUD" then
      EMWRITEScreen "X", 9, 3
	  ElseIf elig_cbud_check = "CBUD" then
		EMWriteScreen "X", 13, 43
	  End If
      transmit
	  row = 1
	  col = 1
	  EMSearch "Ref Nbr: "&searchable_client_name, row, col
	  If row <> 0 then
	    EMWriteScreen "X", row + 1, col - 1 
		transmit
	    call income_type_converter(unearned_type_1)
		EMWriteScreen "__", 8, 8
	    EMWriteScreen "___________", 8, 43
	    EMWriteScreen "_", 8, 58
	    EMWriteScreen converted_income_type, 8, 8
        EMWriteScreen unearned_value_1, 8, 43
        EMWriteScreen unearned_exclusion_1, 8, 58
        transmit
        PF3	  
	  End If
	  PF3	  
	End If
	If unearned_value_2 <> "" then
  	EMReadScreen elig_abud_check, 4, 3, 47
	    elig_abud_check = trim(elig_abud_check)
    EMReadScreen elig_bbud_check, 4, 3, 47  
      elig_bbud_check = trim(elig_bbud_check)
	  EMReadScreen elig_cbud_check, 4, 3, 54
	    elig_cbud_check = trim(elig_cbud_check)
	  If elig_abud_check = "ABUD" then 
	    EMWriteScreen "X", 13, 43
    ElseIf elig_bbud_check = "BBUD" then
      EMWRITEScreen "X", 9, 3
	  ElseIf elig_cbud_check = "CBUD" then
		EMWriteScreen "X", 13, 43
	  End If
      transmit
	  row = 1
	  col = 1
	  EMSearch "Ref Nbr: "&searchable_client_name, row, col
	  If row <> 0 then
	    EMWriteScreen "X", row + 1, col - 1 
		transmit
	    call income_type_converter(unearned_type_2)
		EMWriteScreen "__", 9, 8
	    EMWriteScreen "___________", 9, 43
	    EMWriteScreen "_", 9, 58
	    EMWriteScreen converted_income_type, 9, 8
        EMWriteScreen unearned_value_2, 9, 43
        EMWriteScreen unearned_exclusion_2, 9, 58
        transmit
        PF3	  
	  End If
	  PF3	  
	End If
	If unearned_value_3 <> "" then
  	EMReadScreen elig_abud_check, 4, 3, 47
	    elig_abud_check = trim(elig_abud_check)
    EMReadScreen elig_bbud_check, 4, 3, 47  
      elig_bbud_check = trim(elig_bbud_check)
	  EMReadScreen elig_cbud_check, 4, 3, 54
	    elig_cbud_check = trim(elig_cbud_check)
	  If elig_abud_check = "ABUD" then 
	    EMWriteScreen "X", 13, 43
    ElseIf elig_bbud_check = "BBUD" then
      EMWRITEScreen "X", 9, 3
	  ElseIf elig_cbud_check = "CBUD" then
		EMWriteScreen "X", 13, 43
	  End If
      transmit
	  row = 1
	  col = 1
	  EMSearch "Ref Nbr: "&searchable_client_name, row, col
	  If row <> 0 then
	    EMWriteScreen "X", row + 1, col - 1 
		transmit
	    call income_type_converter(unearned_type_3)
		EMWriteScreen "__", 10, 8
	    EMWriteScreen "___________", 10, 43
	    EMWriteScreen "_", 10, 58
	    EMWriteScreen converted_income_type, 10, 8
        EMWriteScreen unearned_value_3, 10, 43
        EMWriteScreen unearned_exclusion_3, 10, 58
        transmit
        PF3	  
	  End If
	  PF3	  
	End If	    
  End If
End Function

function spec_xfer(bank_number)
	bank_number = trim(bank_number)
	call back_to_self
	EMWriteScreen "SPEC", 16, 43
	EMWriteScreen "XFER", 21, 70
	EMWriteScreen "________", 18, 43
	EMWriteScreen case_number, 18, 43
	transmit
	EMWriteScreen "X", 7, 16
	transmit
	EMReadScreen current_servicing_worker, 7, 18, 61
	current_servicing_worker = trim(current_servicing_worker)
	If current_servicing_worker <> bank_number then 
		PF9
		EMWriteScreen bank_number, 18, 61
		transmit
		call back_to_self
	ElseIf current_servicing_worker = bank_number then
		call back_to_self
	End If
end function

Function auto_pass_all()
	Do
		EMReadScreen assets_test				,	6	,	7	,	5
		EMReadScreen cooperation_test			,	6	,	10	,	5
		EMReadScreen fail_to_file_test			,	6	,	14	,	5
		EMReadScreen obligation_one_month_test	,	6	,	9	,	46
		EMReadScreen verification_test			,	6	,	14	,	46		
		EMReadScreen absence_test				,	6	,	6	,	5
		EMReadScreen assistance_unit_memb_test	,	6	,	8	,	5
		EMReadScreen citizenship_test			,	6	,	9	,	5
		EMReadScreen correctional_facility_test	,	6	,	11	,	5
		EMReadScreen death_test					,	6	,	12	,	5
		EMReadScreen elig_for_other_program_test,	6	,	13	,	5
		EMReadScreen imd_test					,	6	,	15	,	5		
		EMReadScreen income_bdgt_test			,	6	,	6	,	46
		EMReadScreen medicare_elig_test			,	6	,	7	,	46
		EMReadScreen mnsure_system_test			,	6	,	8	,	46
		EMReadScreen obligation_six_month_test	,	6	,	10	,	46
		EMReadScreen other_health_insa_test		,	6	,	11	,	46
		EMReadScreen pare_steppare_test			,	6	,	12	,	46
		EMReadScreen residence_test				,	6	,	13	,	46
		EMReadScreen client_withdrawal_test		,	6	,	15	,	46	
		If assets_test = "FAILED" then
			MsgBox "Please Review the assets test before pressing OK.", "Asset Test Failed"
		End If	
		If cooperation_test = "FAILED" then
			EMWriteScreen "X", 10, 3
			Transmit
			EMReadScreen pben_coop, 6, 10, 28
			EMReadScreen pact_coop, 6, 11, 28
			EMReadScreen disq_coop, 6, 12, 28
			EMReadScreen abps_coop, 6, 13, 28
			EMReadScreen insa_coop, 6, 14, 28
			EMReadScreen memb_coop, 6, 15, 28
			EMReadScreen acci_coop, 6, 16, 28
			If pben_coop = "FAILED" then
				EMWriteScreen "X", 10, 26
				Transmit
				EMReadScreen cash_pben_coop, 6, 10, 31
				EMReadScreen smrt_pben_coop, 6, 11, 31
				If cash_pben_coop = "FAILED" then 
					EMWriteScreen "PASSED", 10, 31
					Transmit
				End If
				If smrt_pben_coop = "FAILED" then 
					EMWriteScreen "PASSED", 11, 31
					Transmit
				End If
				Transmit
			End If		
			If pact_coop = "FAILED" then 
				EMWriteScreen "PASSED", 11, 28		
				Transmit
			End If
			If disq_coop = "FAILED" then 
				EMWriteScreen "PASSED", 12, 28		
				Transmit
			End If
			If abps_coop = "FAILED" then
				EMWriteScreen "PASSED", 13, 28		
				Transmit
			End If
			If insa_coop = "FAILED" then 
				EMWriteScreen "PASSED", 14, 28		
				Transmit
			End If
			If memb_coop = "FAILED" then 
				EMWriteScreen "PASSED", 15, 28		
				Transmit
			End If
			If acci_coop = "FAILED" then 
				EMWriteScreen "PASSED", 16, 28
				Transmit
			End If
			Transmit
		End If	
		If fail_to_file_test = "FAILED" then
			EMWriteScreen "X", 14, 3
			Transmit
			EMReadScreen monthly_household_report_f2f		, 6, 14, 33
			EMReadScreen six_monthly_inc_renewal_f2f		, 6, 15, 33
			EMReadScreen six_monthly_inc_asst_renewal_f2f	, 6, 16, 33
			EMReadScreen twelve_month_elig_renewal_f2f		, 6, 17, 33
			EMReadScreen tyma_quarterly_review_f2f			, 6, 18, 33
			If monthly_household_report_f2f 		= "FAILED" then 
				EMWriteScreen "PASSED", 14, 33
				Transmit
			End If
			If six_monthly_inc_renewal_f2f 			= "FAILED" then 
				EMWriteScreen "PASSED", 15, 33
				Transmit
			End If
			If six_monthly_inc_asst_renewal_f2f 	= "FAILED" then 
				EMWriteScreen "PASSED", 16, 33
				Transmit
			End If
			If twelve_month_elig_renewal_f2f 		= "FAILED" then 
				EMWriteScreen "PASSED", 17, 33
				Transmit
			End If
			If tyma_quarterly_review_f2f 			= "FAILED" then 
				EMWriteScreen "PASSED", 18, 33
				Transmit
			End If
			Transmit
		End If			
		If verification_test = "FAILED" then
			EMWriteScreen "X", 14, 44
			Transmit
			EMReadScreen acct_verif, 6,  5, 10
			EMReadScreen busi_verif, 6,  6, 10
			EMReadScreen jobs_verif, 6,  7, 10
			EMReadScreen disq_imig_verif, 6,  8, 10
			EMReadScreen lump_verif, 6,  9, 10
			EMReadScreen othr_verif, 6, 10, 10
			EMReadScreen pben_verif, 6, 11, 10
			EMReadScreen preg_verif, 6, 12, 10
			EMReadScreen rbic_verif, 6, 13, 10
			EMReadScreen rest_verif, 6, 14, 10
			EMReadScreen secu_verif, 6, 15, 10
			EMReadScreen spon_verif, 6, 16, 10
			EMReadScreen tran_verif, 6, 17, 10
			EMReadScreen unea_verif, 6, 18, 10
			EMReadScreen disq_cit_verif, 6, 19, 10
			EMReadScreen cars_verif, 6, 20, 10
			If acct_verif	= "FAILED" then 
				EMWriteScreen "PASSED", 5, 10
				Transmit
			End If
			If busi_verif	= "FAILED" then 
				EMWriteScreen "PASSED", 6, 10
				Transmit
			End If
			If jobs_verif	= "FAILED" then 
				EMWriteScreen "PASSED", 7, 10
				Transmit
			End If
			If disq_imig_verif	= "FAILED" then 
				EMWriteScreen "PASSED", 8, 10
				Transmit
			End If
			If lump_verif	= "FAILED" then 
				EMWriteScreen "PASSED", 9, 10
				Transmit
			End If
			If othr_verif	= "FAILED" then 
				EMWriteScreen "PASSED", 10, 10
				Transmit
			End If
			If pben_verif	= "FAILED" then 
				EMWriteScreen "PASSED", 11, 10
				Transmit
			End If
			If preg_verif	= "FAILED" then 
				EMWriteScreen "PASSED", 12, 10
				Transmit
			End If
			If rbic_verif	= "FAILED" then 
				EMWriteScreen "PASSED", 13, 10
				Transmit
			End If
			If rest_verif	= "FAILED" then 
				EMWriteScreen "PASSED", 14, 10
				Transmit
			End If
			If secu_verif	= "FAILED" then 
				EMWriteScreen "PASSED", 15, 10
				Transmit
			End If
			If spon_verif	= "FAILED" then 
				EMWriteScreen "PASSED", 16, 10
				Transmit
			End If
			If tran_verif	= "FAILED" then 
				EMWriteScreen "PASSED", 17, 10
				Transmit
			End If
			If unea_verif	= "FAILED" then 
				EMWriteScreen "PASSED", 18, 10
				Transmit
			End If
			If disq_cit_verif	= "FAILED" then 
				EMWriteScreen "PASSED", 19, 10
				Transmit
			End If
			If cars_verif	= "FAILED" then 
				EMWriteScreen "PASSED", 20, 10
				Transmit
			End If
			Transmit
		End If		
		If absence_test 				= "FAILED" then
			EMWriteScreen "PASSED",  6, 5
			Transmit
		End If
		If assistance_unit_memb_test 	= "FAILED" then
			EMWriteScreen "PASSED",  8, 5
			Transmit
		End If
		If citizenship_test 			= "FAILED" then 
			EMWriteScreen "PASSED",  9, 5
			Transmit
		End If
		If correctional_facility_test 	= "FAILED" then 
			EMWriteScreen "PASSED", 11, 5
			Transmit
		End If
		If death_test 					= "FAILED" then 
			EMWriteScreen "PASSED", 12, 5
			Transmit
		End If
		If elig_for_other_program_test 	= "FAILED" then 
			EMWriteScreen "PASSED", 13, 5
			Transmit
		End If
		If imd_test 					= "FAILED" then 
			EMWriteScreen "PASSED", 15, 5
			Transmit
		End If
		
		If income_bdgt_test 			= "FAILED" then 
			EMWriteScreen "PASSED",  6, 46
			Transmit
		End If
		If medicare_elig_test 			= "FAILED" then	
			EMWriteScreen "PASSED",  7, 46
			Transmit
		End If
		If mnsure_system_test 			= "FAILED" then 
			EMWriteScreen "PASSED",  8, 46
			Transmit
		End If
		If obligation_six_month_test 	= "FAILED" then 
			EMWriteScreen "N/A   ", 10, 46
			Transmit
		End If
		If other_health_insa_test 		= "FAILED" then 
			EMWriteScreen "PASSED", 11, 46
			Transmit
		End If
		If pare_steppare_test 			= "FAILED" then 
			EMWriteScreen "PASSED", 12, 46
			Transmit
		End If
		If residence_test 				= "FAILED" then	
			EMWriteScreen "PASSED", 13, 46
			Transmit
		End If
		If client_withdrawal_test 		= "FAILED" then 
			EMWriteScreen "PASSED", 15, 46
			Transmit
		End If		
		If obligation_one_month_test = "FAILED" then
			EMWriteScreen "N/A   ", 9, 46
			Transmit
		End If
	Loop until assets_test 					<> "FAILED" _
		   and cooperation_test 			<> "FAILED" _
		   and fail_to_file_test 			<> "FAILED" _
		   and obligation_one_month_test 	<> "FAILED" _
		   and verification_test			<> "FAILED" _   
		   and absence_test					<> "FAILED" _
		   and assistance_unit_memb_test	<> "FAILED" _
		   and citizenship_test				<> "FAILED" _
		   and correctional_facility_test	<> "FAILED" _
		   and death_test					<> "FAILED" _
		   and elig_for_other_program_test	<> "FAILED" _
		   and imd_test						<> "FAILED" _	
		   and income_bdgt_test				<> "FAILED" _
		   and medicare_elig_test			<> "FAILED" _
		   and mnsure_system_test			<> "FAILED" _
		   and obligation_six_month_test	<> "FAILED" _
		   and other_health_insa_test		<> "FAILED" _
		   and pare_steppare_test			<> "FAILED" _
		   and residence_test				<> "FAILED" _
		   and client_withdrawal_test		<> "FAILED"
End Function

Function set_person_test(reason, result)
	X = 8
	x_mits = 1
	If 		reason = "Failure To Provide Information" 	or _
			reason = "IEVS" 							or _
			reason = "Medical Support" 					or _
			reason = "Other Health Insurance" 			or _
			reason = "Social Security Number" 			or _
			reason = "Third Party Liability" 			then
		EMWriteScreen "X", 10, 3
		Transmit
		x_mits = 2
		x = 15
	ElseIf 	reason = "Cash" or _
			reason = "SMRT" then
		EMWriteScreen "X", 10, 3
		Transmit
		EMWriteScreen "X", 10, 26
		Transmit
		x_mits = 3
		x = 10
	ElseIf 	reason = "Monthly Household Report" 	or _
			reason = "6 Month Income Renewal" 		or _
			reason = "6 Month Income/Asset Renewal" or _
			reason = "12 Month Eligibility Renewal" or _
			reason = "TYMA Quarterly Review" 		then
		EMWriteScreen "X", 14, 3
		Transmit
		x_mits = 2
	ElseIf 	reason = "Accounts" 					or _
			reason = "Business Income" 				or _
			reason = "Earned Income"				or _
			reason = "Immigration Status" 			or _
			reason = "Lump Sum Income" 				or _
			reason = "Other Assets" 				or _
			reason = "Potential Benefits" 			or _
			reason = "Pregnancy" 					or _
			reason = "Room/Board Income" 			or _
			reason = "Real Estate" 					or _
			reason = "Securities" 					or _
			reason = "Sponsor Income And Assets" 	or _
			reason = "Transferred Assets" 			or _
			reason = "Unearned Income" 				or _
			reason = "12 Month Eligibility Renewal" or _
			reason = "US Citizenship/Identity" 		then
		EMWriteScreen "X", 14, 44
		Transmit
		x_mits = 2
		x = 15
	End If
	row = 1
	col = 1
	runs = 0
	EMSearch reason, row, col
	If row <> 0 then
		EMWriteScreen result, row, col - X
		Do 
			Transmit
			runs = runs + 1
		Loop until runs = x_mits
	End If	
End Function

Function mnsure_fail_person_test	
	BeginDialog person_test_fail, 0, 0, 149, 62, "Person Test Fail"
		ButtonGroup ButtonPressed
      PushButton 123, 45, 22, 12, "Back", back_button
      PushButton 91, 45, 29, 12, "Submit", submit_button
    CheckBox 7, 14, 10, 10, "", PASSED_failure_to_provide_information
		CheckBox 26, 14, 118, 10, "     Failure To Provide Information", FAILED_failure_to_provide_information
		CheckBox 7, 27, 10, 10, "", PASSED_withdrawal_client_request
		CheckBox 26, 27, 110, 10, "     Withdrawal/Client Request", FAILED_withdrawal_client_request
		Text 3, 3, 17, 8, "Pass"
		Text 24, 3, 12, 8, "Fail"
	EndDialog
	Do
		error_count = 0
		Dialog person_test_fail	
			If buttonpressed = back_button then Exit Function
		If (PASSED_failure_to_provide_information = 1 and FAILED_failure_to_provide_information = 1) or _
		   (PASSED_withdrawal_client_request = 1 	  and FAILED_withdrawal_client_request = 1)	then 
				MsgBox "You cannot pass and fail a test at the same time. Please check your selection and try again.","Selection Error"
				error_count = error_count + 1
		End If
	Loop until error_count = 0
	test_1 = ""
	test_2 = ""
	If  PASSED_failure_to_provide_information = 1 then	test_1 = "Passed"
	If  FAILED_failure_to_provide_information = 1 then	test_1 = "Failed"
	If  PASSED_withdrawal_client_request = 1 then	test_2 = "Passed"
	If  FAILED_withdrawal_client_request = 1 then	test_2 = "Failed"
	If test_1 <> "" or test_2 <> "" then
		A = 0
		col = 17
		For i = 6 to 1 Step -1	
			EMWriteScreen "X", 7, col
			EMReadScreen B, 1, 7, col
			If B = "X" then A = A + 1
			col = col + 11
		Next
		For i = A To 1 Step -1
			Transmit
			If test_1 <> "" then call set_person_test("Failure To Provide Information", test_1)
			If test_2 <> "" then call set_person_test("Withdrawal/Client Request", test_2)
		Next
		Transmit
	End If
End Function

Function enter_income_information(A)
	If A = 1 then 
		V = client_01_name
		W = set_included_income_01
		X = hh_count_01_searchable
		Y = unearned_income_01
		Z = earned_income_01
	ElseIf A = 2 then
		V = client_02_name
		W = set_included_income_02
		X = hh_count_02_searchable
		Y = unearned_income_02
		Z = earned_income_02
	ElseIf A = 3 then
		V = client_03_name
		W = set_included_income_03
		X = hh_count_03_searchable
		Y = unearned_income_03
		Z = earned_income_03
	ElseIf A = 4 then
		V = client_04_name
		W = set_included_income_04
		X = hh_count_04_searchable
		Y = unearned_income_04
		Z = earned_income_04
	ElseIf A = 5 then
		V = client_05_name
		W = set_included_income_05
		X = hh_count_05_searchable
		Y = unearned_income_05
		Z = earned_income_05
	ElseIf A = 6 then
		V = client_06_name
		W = set_included_income_06
		X = hh_count_06_searchable
		Y = unearned_income_06
		Z = earned_income_06
	ElseIf A = 7 then
		V = client_07_name
		W = set_included_income_07
		X = hh_count_07_searchable
		Y = unearned_income_07
		Z = earned_income_07
	ElseIf A = 8 then
		V = client_08_name
		W = set_included_income_08
		X = hh_count_08_searchable
		Y = unearned_income_08
		Z = earned_income_08
	ElseIf A = 9 then
		V = client_09_name
		W = set_included_income_09
		X = hh_count_09_searchable
		Y = unearned_income_09
		Z = earned_income_09
	ElseIf A = 10 then
		V = client_10_name
		W = set_included_income_10
		X = hh_count_10_searchable
		Y = unearned_income_10
		Z = earned_income_10
	ElseIf A = 11 then
		V = client_11_name
		W = set_included_income_11
		X = hh_count_11_searchable
		Y = unearned_income_11
		Z = earned_income_11
	ElseIf A = 12 then
		V = client_12_name
		W = set_included_income_12
		X = hh_count_12_searchable
		Y = unearned_income_12
		Z = earned_income_12
	End If	
	If V <> "" and W = "Y" then 
		If Y = 1 or Z = 1 then
			call additional_income_info(V,Y,Z)
				If buttonpressed = 0 then stopscript
			call enter_income_screen_information(X,V,current_client_selected,earned_type_1,earned_type_2,earned_type_3,unearned_type_1,unearned_type_2,unearned_type_3,earned_value_1,earned_value_2,earned_value_3,unearned_value_1,unearned_value_2,unearned_value_3,earned_exclusion_1,earned_exclusion_2,earned_exclusion_3,unearned_exclusion_1,unearned_exclusion_2,unearned_exclusion_3)
		End If
	End If
End Function

Function create_or_edit_panel(A,B)
	'A = Panel Type
	'B = HH_member
	If B = "n/a" then B = ""
	If A = "case" then
		EMReadScreen panel_no_exist, 28, 24, 7
			panel_no_exist = trim(panel_no_exist)
		If panel_no_exist = "DOES NOT EXIST FOR THIS CASE" then
		EMWriteScreen "NN", 20, 79
			transmit
		Elseif panel_no_exist = "" then
			PF9
		End If
	ElseIf A = "person" then
		EMReadScreen panel_already_exist, 1, 2, 78
		If panel_already_exist = "1" then 
			PF9
		ElseIf panel_already_exist = "0" then
			EMWriteScreen B, 20, 76
			EMWriteScreen "NN", 20, 79
			transmit
		End If
	End If
End Function

Function screen_error_check()
	EMReadScreen screen_errors, 79, 24, 1
	screen_errors = trim(screen_errors)
	If screen_errors <> "" then
		MsgBox "There was an error, warning, or other notification message found on this screen please review this and make corrections as necessary before you press OK to continue." , "Error, Warning, or Other Notice Found"
	End If
End Function
