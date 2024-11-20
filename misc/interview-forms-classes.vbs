Set xmlDoc = CreateObject("Microsoft.XMLDOM")

class form_questions
	public number
	public dialog_phrasing
	public note_phrasing
	public doc_phrasing

	public sub_number
	public sub_phrase
	public sub_note_phrase
	public sub_answer

	public info_type
	public caf_answer
	public answer_is_array
	public make_array_checkboxes
	public write_in_info
	public interview_notes
	public item_info_list
	public item_note_info_list
	public item_ans_list
	public item_detail_list
	public allow_prefil
	public supplemental_questions
	public entirely_blank
	public associated_array

	public detail_array_exists
	public detail_source
	public detail_interview_notes
	public detail_write_in_info
	public detail_verif_status
	public detail_verif_notes
	public detail_edit_btn

	public detail_resident_name
	public detail_value
	public detail_type
	public detail_hourly_wage
	public detail_hours_per_week
	public detail_business
	public detail_monthly_amount
	public detail_date
	public detail_frequency
	public detail_amount
	public detail_current
	public detail_explain
	public detail_button_label

	public housing_payment
	public heat_air_checkbox
	public electric_checkbox
	public phone_checkbox
	public subsidy_yn
	public subsidy_amount

	public verif_status
	public verif_notes
	public verif_verbiage

	public guide_btn
	public verif_btn
	public prefil_btn
	public add_to_array_btn
	public edit_in_array_btn

	public mandated
	public error_info
	public error_verbiage

	public dialog_page_numb
	public dialog_order
	public dialog_height

	public sub add_detail_item(index_add)
		ReDim Preserve detail_interview_notes(index_add)
		ReDim Preserve detail_write_in_info(index_add)
		ReDim Preserve detail_verif_status(index_add)
		ReDim Preserve detail_verif_notes(index_add)
		ReDim Preserve detail_edit_btn(index_add)
		If IsArray(detail_type) 			Then ReDim Preserve detail_type(index_add)
		If IsArray(detail_resident_name) 	Then ReDim Preserve detail_resident_name(index_add)
		If IsArray(detail_value) 			Then ReDim Preserve detail_value(index_add)
		If IsArray(detail_hourly_wage) 		Then ReDim Preserve detail_hourly_wage(index_add)
		If IsArray(detail_hours_per_week) 	Then ReDim Preserve detail_hours_per_week(index_add)
		If IsArray(detail_business) 		Then ReDim Preserve detail_business(index_add)
		If IsArray(detail_monthly_amount) 	Then ReDim Preserve detail_monthly_amount(index_add)
		If IsArray(detail_date) 			Then ReDim Preserve detail_date(index_add)
		If IsArray(detail_frequency) 		Then ReDim Preserve detail_frequency(index_add)
		If IsArray(detail_amount) 			Then ReDim Preserve detail_amount(index_add)
		If IsArray(detail_current) 			Then ReDim Preserve detail_current(index_add)
		If IsArray(detail_explain) 			Then ReDim Preserve detail_explain(index_add)
	end sub

	public sub display_in_dialog(y_pos, question_yn, question_notes, question_interview_notes, addtl_question, TEMP_ARRAY)
		question_answers = ""+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Blank"
		If InStr(dialog_phrasing, "For MA-LTC") <> 0 Then question_answers = question_answers+chr(9)+"NoLTC"
		' MsgBox number & "." & dialog_phrasing & vbCr &_
		' 		"question_yn - ~" & question_yn & "~" & vbCr &_
		' 		"question_notes - " & question_notes & vbCr &_
		' 		"question_interview_notes - " & question_interview_notes & vbCr &_
		' 		"addtl_question - " & addtl_question
		If answer_is_array = false Then question_yn = caf_answer
		If answer_is_array = true Then
			If info_type <> "unea" Then
				For i = 0 to UBound(item_ans_list)
					TEMP_ARRAY(i) = item_ans_list(i)
				Next
			End If
			If info_type = "unea" Then
				unea_1_yn  = item_ans_list(0)
				unea_1_amt = item_detail_list(0)
				unea_2_yn  = item_ans_list(1)
				unea_2_amt = item_detail_list(1)
				unea_3_yn  = item_ans_list(2)
				unea_3_amt = item_detail_list(2)
				unea_4_yn  = item_ans_list(3)
				unea_4_amt = item_detail_list(3)
				unea_5_yn  = item_ans_list(4)
				unea_5_amt = item_detail_list(4)
				unea_6_yn  = item_ans_list(5)
				unea_6_amt = item_detail_list(5)
				unea_7_yn  = item_ans_list(6)
				unea_7_amt = item_detail_list(6)
				unea_8_yn  = item_ans_list(7)
				unea_8_amt = item_detail_list(7)
				unea_9_yn  = item_ans_list(8)
				unea_9_amt = item_detail_list(8)
			End If
		End If
		addtl_question = sub_answer
		question_notes = write_in_info
		question_interview_notes = interview_notes

		If detail_array_exists = true Then
			grp_len = 35
			' MsgBox " is detail_interview_notes an array? - " & IsArray(detail_interview_notes) & vbCr & number
			for each_item = 0 to UBOUND(detail_interview_notes)
				show_info = False
				If detail_source = "jobs" Then
					' MsgBox "each_item - " & each_item & vbCr & IsArray(detail_business)
					If detail_business(each_item) <> "" Then show_info = True
					If detail_resident_name(each_item) <> "" Then show_info = True
					If detail_monthly_amount(each_item) <> "" Then show_info = True
					If detail_hourly_wage(each_item) <> "" Then show_info = True
				ElseIf detail_source = "assets" Then
					If detail_resident_name(each_item) <> "" Then show_info = True
					If detail_type(each_item) <> "" Then show_info = True
					If detail_value(each_item) <> "" Then show_info = True
					If detail_explain(each_item) <> "" Then show_info = True
				ElseIf detail_source = "unea" Then
					If detail_resident_name(each_item) <> "" Then show_info = True
					If detail_type(each_item) <> "" Then show_info = True
					If detail_date(each_item) <> "" Then show_info = True
					If detail_amount(each_item) <> "" Then show_info = True
					If detail_frequency(each_item) <> "" Then show_info = True
				ElseIf detail_source = "shel-hest" Then
					If detail_type(each_item) <> "" Then show_info = True
					If detail_amount(each_item) <> "" Then show_info = True
					If detail_frequency(each_item) <> "" Then show_info = True
				ElseIf detail_source = "expense" Then
					If detail_resident_name(each_item) <> "" Then show_info = True
					If detail_amount(each_item) <> "" Then show_info = True
					If detail_current(each_item) <> "" Then show_info = True
				ElseIf detail_source = "winnings" Then
					If detail_resident_name(each_item) <> "" Then show_info = True
					If detail_amount(each_item) <> "" Then show_info = True
					If detail_date(each_item) <> "" Then show_info = True
				ElseIf detail_source = "changes" Then
					If detail_resident_name(each_item) <> "" Then show_info = True
					If detail_date(each_item) <> "" Then show_info = True
					If detail_explain(each_item) <> "" Then show_info = True
				End If
				If show_info = True then grp_len = grp_len + 20
			next
			If detail_source = "shel-hest" Then grp_len = grp_len + 20

			GroupBox 5, y_pos, 475, grp_len, number & "." & dialog_phrasing
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_yn
			Text 95, y_pos, 25, 10, "write-in:"
			EditBox 120, y_pos - 5, 350, 15, question_notes
			PushButton 425, y_pos-20, 55, 10, detail_button_label, add_to_array_btn

			first_item = TRUE
			for each_item = 0 to UBOUND(detail_interview_notes)
				If detail_source = "jobs" Then
					If detail_business(each_item) <> "" OR detail_resident_name(each_item)<> "" OR detail_monthly_amount(each_item) <> "" OR detail_hourly_wage(each_item) <> "" OR detail_hours_per_week(each_item) <> "" Then
						If first_item = TRUE Then y_pos = y_pos + 20
						first_item = FALSE
						If detail_verif_status(each_item) = "" Then Text 15, y_pos, 395, 10, "Employer: " & detail_business(each_item) & "  - Employee: " & detail_resident_name(each_item) & "   - Gross Monthly Earnings: $ " & detail_monthly_amount(each_item)
						If detail_verif_status(each_item) <> "" Then Text 15, y_pos, 395, 10, "Employer: " & detail_business(each_item) & "  - Employee: " & detail_resident_name(each_item) & "   - Gross Monthly Earnings: $ " & detail_monthly_amount(each_item) & "   - Verification - " & detail_verif_status(each_item)
						PushButton 450, y_pos, 20, 10, "EDIT", detail_edit_btn(each_item)
						y_pos = y_pos + 10
					End If

				ElseIf detail_source = "assets" Then
					If detail_resident_name(each_item) <> "" OR detail_type(each_item) <> "" OR detail_value(each_item) <> "" Then
						If first_item = TRUE Then y_pos = y_pos + 20
						first_item = FALSE
						If detail_verif_status(each_item) = "" Then Text 15, y_pos, 395, 10, "Owner: " & detail_resident_name(each_item) & "  - Type: " & detail_type(each_item) & "  - Value: $ " & detail_value(each_item)
						If detail_verif_status(each_item) <> "" Then Text 15, y_pos, 395, 10, "Owner: " & detail_resident_name(each_item) & "  - Type: " & detail_type(each_item) & "  - Value: $ " & detail_value(each_item) & "   - Verification - " & detail_verif_status(each_item)
						PushButton 450, y_pos, 20, 10, "EDIT", detail_edit_btn(each_item)
						y_pos = y_pos + 10
					End If
				ElseIf detail_source = "unea" Then
					If detail_resident_name(each_item) <> "" OR detail_type(each_item) <> "" OR detail_date(each_item) <> "" OR detail_amount(each_item) <> "" OR detail_frequency(each_item) <> "" Then
						If first_item = TRUE Then y_pos = y_pos + 20
						first_item = FALSE
						If detail_verif_status(each_item) = "" Then Text 15, y_pos, 395, 10, "Name: " & detail_resident_name(each_item) & "  - Type: " & detail_type(each_item) & "   - Start Date: $ " & detail_date(each_item) & "   - Amount: $ " & detail_amount(each_item) & "   - Freq.: " & detail_frequency(each_item)
						If detail_verif_status(each_item) <> "" Then Text 15, y_pos, 395, 10, "Name: " & detail_resident_name(each_item) & "  - Type: " & detail_type(each_item) & "   - Start Date: $ " & detail_date(each_item) & "   - Amount: $ " & detail_amount(each_item) & "   - Freq.: " & detail_frequency(each_item) & "   - Verification - " & detail_verif_status(each_item)

						PushButton 450, y_pos, 20, 10, "EDIT", detail_edit_btn(each_item)
						y_pos = y_pos + 10
					End If
				ElseIf detail_source = "shel-hest" Then
					If detail_type(each_item) <> "" OR  detail_amount(each_item) <> "" OR detail_frequency(each_item) <> "" Then
						If first_item = TRUE Then y_pos = y_pos + 20
						first_item = FALSE
						If detail_verif_status(each_item) = "" Then Text 15, y_pos, 395, 10, "Type: " & detail_type(each_item) & "  - Amount: $ " & detail_amount(each_item) & "  - Frequency: " & detail_frequency(each_item)
						If detail_verif_status(each_item) <> "" Then Text 15, y_pos, 395, 10, "Type: " & detail_type(each_item) & "  - Amount: $ " & detail_amount(each_item) & "  - Frequency: " & detail_frequency(each_item) & "   - Verification - " & detail_verif_status(each_item)
						PushButton 450, y_pos, 20, 10, "EDIT", detail_edit_btn(each_item)
						y_pos = y_pos + 10
					End If
				ElseIf detail_source = "expense" Then
					If detail_resident_name(each_item) <> "" OR detail_amount(each_item) <> "" OR detail_current(each_item) <> "" Then
						If first_item = TRUE Then y_pos = y_pos + 20
						first_item = FALSE
						If detail_verif_status(each_item) = "" Then Text 15, y_pos, 395, 10, "Payer: " & detail_resident_name(each_item) & "  - Amount: $ " & detail_amount(each_item) & "  - Currently Paying: " & detail_current(each_item)
						If detail_verif_status(each_item) <> "" Then Text 15, y_pos, 395, 10, "Payer: " & detail_resident_name(each_item) & "  - Amount: $ " & detail_amount(each_item) & "  - Currently Paying: " & detail_current(each_item) & "   - Verification - " & detail_verif_status(each_item)
						PushButton 450, y_pos, 20, 10, "EDIT", detail_edit_btn(each_item)
						y_pos = y_pos + 10
					End If
				ElseIf detail_source = "winnings" Then
					If detail_resident_name(each_item) <> "" or detail_amount(each_item) <> "" or detail_date(each_item) <> "" Then
						If first_item = TRUE Then y_pos = y_pos + 20
						first_item = FALSE
						If detail_verif_status(each_item) = "" Then Text 15, y_pos, 395, 10, "Winner: " & detail_resident_name(each_item) & "  - Amount: $ " & detail_amount(each_item) & "  - Date of Win: " & detail_date(each_item)
						If detail_verif_status(each_item) <> "" Then Text 15, y_pos, 395, 10, "Winner: " & detail_resident_name(each_item) & "  - Amount: $ " & detail_amount(each_item) & "  - Date of Win: " & detail_date(each_item) & "   - Verification - " & detail_verif_status(each_item)
						PushButton 450, y_pos, 20, 10, "EDIT", detail_edit_btn(each_item)
						y_pos = y_pos + 10
					End If
				ElseIf detail_source = "changes" Then
					If detail_resident_name(each_item) <> "" OR detail_date(each_item) <> "" OR detail_explain(each_item) <> "" Then
						If first_item = TRUE Then y_pos = y_pos + 20
						first_item = FALSE
						If detail_verif_status(each_item) = "" Then Text 15, y_pos, 395, 10, "Who: " & detail_resident_name(each_item) & "  - Date: " & detail_date(each_item) & "  - Explain: " & detail_explain(each_item)
						If detail_verif_status(each_item) <> "" Then Text 15, y_pos, 395, 10, "Who: " & detail_resident_name(each_item) & "  - Date: " & detail_date(each_item) & "  - Explain: " & detail_explain(each_item) & "   - Verification - " & detail_verif_status(each_item)
						PushButton 450, y_pos, 20, 10, "EDIT", detail_edit_btn(each_item)
						y_pos = y_pos + 10
					End If
				End If
			next
			If first_item = TRUE Then y_pos = y_pos + 10

			If detail_source = "shel-hest" Then
				housing_info_txt = "Housing Payment: $ " & housing_payment
				If heat_air_checkbox = unchecked Then housing_info_txt = housing_info_txt & "  -   [ ] Heat/AC "
				If heat_air_checkbox = checked Then housing_info_txt = housing_info_txt & "  -   [X] Heat/AC "
				If electric_checkbox = unchecked Then housing_info_txt = housing_info_txt & "   [ ] Electric "
				If electric_checkbox = checked Then housing_info_txt = housing_info_txt & "   [X] Electric "
				If phone_checkbox = unchecked Then housing_info_txt = housing_info_txt & "   [ ] Phone "
				If phone_checkbox = checked Then housing_info_txt = housing_info_txt & "   [X] Phone "

				subsidy_info_tx = "Subsidy: " & subsidy_yn & "    Subsidy Amount: $ " & subsidy_amount

				Text 10, y_pos, 395, 10, housing_info_txt
				Text 10, y_pos+10, 395, 10, subsidy_info_tx
				y_pos = y_pos +20
			End If

			y_pos = y_pos + 15

		ElseIf info_type = "standard" Then
			'funcitonality here
			GroupBox 5, y_pos, 475, dialog_height-5, number & "." & dialog_phrasing
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_yn
			Text 95, y_pos, 25, 10, "write-in:"
			If verif_status = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_notes
			Else
				EditBox 120, y_pos - 5, 235, 15, question_notes
				Text 360, y_pos, 110, 10, "Q" & number & " - Verification - " & verif_status
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", verif_btn
			y_pos = y_pos + 20
		ElseIf info_type = "unea" Then
			GroupBox 5, y_pos, 475, dialog_height-5, number & "." & dialog_phrasing
			PushButton 385, y_pos + 5, 90, 13, "ALL Q. " & number & " Answered 'No'", prefil_btn
			y_pos = y_pos + 20

			col_1_1 = 15
			col_1_2 = 55
			col_1_3 = 115

			col_2_1 = 165
			col_2_2 = 205
			col_2_3 = 260

			col_3_1 = 320
			col_3_2 = 360
			col_3_3 = 430

			Text 	col_1_1, 		y_pos, 40, 10, "CAF Answer"
			Text 	col_1_3 - 3, 	y_pos, 40, 10, "CAF Amount"
			Text 	col_2_1, 		y_pos, 40, 10, "CAF Answer"
			Text 	col_2_3 - 3, 	y_pos, 40, 10, "CAF Amount"
			Text 	col_3_1, 		y_pos, 40, 10, "CAF Answer"
			Text 	col_3_3 - 3, 	y_pos, 40, 10, "CAF Amount"
			y_pos = y_pos + 15

			DropListBox 	col_1_1, 	y_pos, 		35, 45, question_answers, unea_1_yn
			Text 			col_1_2, 	y_pos + 5, 	60, 10, item_info_list(0) & "                  $"
			EditBox 		col_1_3,	y_pos, 		35, 15, unea_1_amt
			DropListBox 	col_2_1, 	y_pos, 		35, 45, question_answers, unea_2_yn
			Text 			col_2_2, 	y_pos + 5, 	60, 10, item_info_list(1) & "                $"
			EditBox 		col_2_3, 	y_pos, 		35, 15, unea_2_amt
			DropListBox 	col_3_1, 	y_pos, 		35, 45, question_answers, unea_3_yn
			Text 			col_3_2, 	y_pos + 5, 	70, 10, item_info_list(2) & "                          $"
			EditBox 		col_3_3, 	y_pos, 		35, 15, unea_3_amt
			y_pos = y_pos + 15

			DropListBox 	col_1_1, 	y_pos, 		35, 45, question_answers, unea_4_yn
			Text 			col_1_2, 	y_pos + 5, 	60, 10, item_info_list(3) & "                       $"
			EditBox 		col_1_3, 	y_pos, 		35, 15, unea_4_amt
			DropListBox 	col_2_1, 	y_pos, 		35, 45, question_answers, unea_5_yn
			Text 			col_2_2, 	y_pos + 5, 	60, 10, item_info_list(4) & "                $"
			EditBox 		col_2_3, 	y_pos, 		35, 15, unea_5_amt
			DropListBox 	col_3_1, 	y_pos, 		35, 45, question_answers, unea_6_yn
			Text 			col_3_2, 	y_pos + 5, 	85, 10, item_info_list(5) & "     $"
			EditBox 		col_3_3, 	y_pos, 		35, 15, unea_6_amt
			y_pos = y_pos + 15

			DropListBox 	col_1_1, 	y_pos, 		35, 45, question_answers, unea_7_yn
			Text 			col_1_2, 	y_pos + 5, 	60, 10, item_info_list(6) & "  $"
			EditBox 		col_1_3, 	y_pos, 		35, 15, unea_7_amt
			DropListBox 	col_2_1, 	y_pos, 		35, 45, question_answers, unea_8_yn
			Text 			col_2_2, 	y_pos + 5, 	60, 10, item_info_list(7) & "             $"
			EditBox 		col_2_3,	y_pos, 		35, 15, unea_8_amt
			DropListBox 	col_3_1, 	y_pos, 		35, 45, question_answers, unea_9_yn
			Text 			col_3_2, 	y_pos + 5, 	110, 10,item_info_list(8) & "       $"
			EditBox 		col_3_3, 	y_pos, 		35, 15, unea_9_amt
			y_pos = y_pos + 25

			Text 15, y_pos, 25, 10, "Write-in:"
			If verif_status = "" Then
				EditBox 40, y_pos - 5, 435, 15, question_notes
			Else
				EditBox 40, y_pos - 5, 315, 15, question_notes
				Text 360, y_pos, 110, 10, "Q" & number & " - Verification - " & verif_status
			End If
			' Text 360, y_pos, 105, 10, "Q10 - Verification - " & question_10_verif_yn
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", verif_btn
			y_pos = y_pos + 25



		ElseIf info_type = "housing" or info_type = "utilities" or info_type = "assets" or info_type = "msa" or info_type = "stwk" Then


			GroupBox 5, y_pos, 475, dialog_height-5, number & "." & dialog_phrasing
			If allow_prefil = true Then PushButton 385, y_pos+5, 90, 13, "ALL Q. " & number & " Answered 'No'", prefil_btn
			y_pos = y_pos + 15

			col_1_1 = ""
			col_1_2 = ""
			col_2_1 = ""
			col_2_2 = ""
			col_3_1 = ""
			col_3_2 = ""
			If info_type = "housing" Then
				col_1_1 = 15
				col_1_2 = 85
				col_2_1 = 220
				col_2_2 = 290
				drplst_len = 60
				txt_len = 135
			End If
			If info_type = "utilities" Then
				col_1_1 = 20
				col_1_2 = 65
				col_2_1 = 185
				col_2_2 = 230
				col_3_1 = 335
				col_3_2 = 380
				drplst_len = 35
				txt_len = 80
			End If
			If info_type = "assets" Then
				col_1_1 = 25
				col_1_2 = 90
				col_2_1 = 230
				col_2_2 = 295
				drplst_len = 60
				txt_len = 175
			End If
			If info_type = "msa" or info_type = "stwk" Then
				col_1_1 = 25
				col_1_2 = 90
				col_2_1 = 230
				col_2_2 = 295
				drplst_len = 60
				txt_len = 140
			End If


			If make_array_checkboxes = true Then
				If info_type = "msa" Then Text 	col_1_1, y_pos, 200, 10, "CAF Answer (check the expenses indicated on the CAF)"
				If info_type = "stwk" Then Text col_1_1, y_pos, 200, 10, "CAF Answer (check the answers indicated on the CAF)"
				y_pos = y_pos + 15

				CheckBox  	col_1_1, y_pos, drplst_len+txt_len, 10, item_info_list(0), TEMP_ARRAY(0)
				CheckBox  	col_2_1, y_pos, drplst_len+txt_len, 10, item_info_list(1), TEMP_ARRAY(1)
				i = 2
				y_pos = y_pos + 10
				CheckBox  	col_1_1, y_pos, drplst_len+txt_len, 10, item_info_list(i), TEMP_ARRAY(i)
				CheckBox  	col_2_1, y_pos, drplst_len+txt_len, 10, item_info_list(i+1), TEMP_ARRAY(i+1)
				i = i + 2
				y_pos = y_pos + 10
				' CheckBox  	col_1_1, y_pos, drplst_len+txt_len, 45, item_info_list(i), TEMP_ARRAY(i)
				' CheckBox  	col_2_1, y_pos, drplst_len+txt_len, 45, item_info_list(i+1), TEMP_ARRAY(i+1)
				' i = i + 2
				y_pos = y_pos + 5
			Else

				Text 	col_1_1, 		y_pos, 40, 10, "CAF Answer"
				Text 	col_2_1, 		y_pos, 40, 10, "CAF Answer"
				If col_3_1 <> "" Then Text 	col_3_1, y_pos, 40, 10, "CAF Answer"
				y_pos = y_pos + 15

				DropListBox 	col_1_1, y_pos - 5, drplst_len, 45, question_answers, TEMP_ARRAY(0)
				Text 			col_1_2, y_pos, 	txt_len, 	10, item_info_list(0)
				DropListBox 	col_2_1, y_pos - 5, drplst_len, 45, question_answers, TEMP_ARRAY(1)
				Text 			col_2_2, y_pos, 	txt_len, 	10, item_info_list(1)
				i = 2
				If col_3_1 <> "" Then
					DropListBox 	col_3_1, y_pos - 5, drplst_len, 45, question_answers, TEMP_ARRAY(i)
					Text 			col_3_2, y_pos, 	txt_len, 	10, item_info_list(i)
					i = i + 1
				End If
				y_pos = y_pos + 15

				DropListBox 	col_1_1, y_pos - 5, drplst_len, 45, question_answers, TEMP_ARRAY(i)
				Text 			col_1_2, y_pos, 	txt_len, 	10, item_info_list(i)
				DropListBox 	col_2_1, y_pos - 5, drplst_len, 45, question_answers, TEMP_ARRAY(i+1)
				Text 			col_2_2, y_pos, 	txt_len, 	10, item_info_list(i+1)
				i = i + 2
				If col_3_1 <> "" Then
					DropListBox 	col_3_1, y_pos - 5, drplst_len, 45, question_answers, TEMP_ARRAY(i)
					Text 			col_3_2, y_pos, 	txt_len, 	10, item_info_list(i)
					i = i + 1
				End If
				y_pos = y_pos + 15

				If info_type = "utilities" Then txt_len = 375
				increase_y_pos = false
				If i =< UBound(TEMP_ARRAY) Then
					DropListBox 	col_1_1, y_pos - 5, drplst_len, 45, question_answers, TEMP_ARRAY(i)
					Text 			col_1_2, y_pos, 	txt_len, 	10, item_info_list(i)
					i = i + 1
					increase_y_pos = true
				End If
				If i =< UBound(TEMP_ARRAY) Then
					DropListBox 	col_2_1, y_pos - 5, drplst_len, 45, question_answers, TEMP_ARRAY(5)
					Text 			col_2_2, y_pos, 	txt_len, 	10, item_info_list(5)
					i = i + 1
					increase_y_pos = true
				End If
				If increase_y_pos = true Then y_pos = y_pos + 15

				If i =< UBound(TEMP_ARRAY) Then
					DropListBox 	col_1_1, y_pos - 5, drplst_len, 45, question_answers, TEMP_ARRAY(6)
					Text 			col_1_2, y_pos, 	txt_len, 	10, item_info_list(6)
					i = i + 1
					y_pos = y_pos + 15
				End If
			End If
			y_pos = y_pos + 5
			If sub_phrase <> "" Then
				y_pos = y_pos - 5

				phrase_width = len(sub_number & "." & sub_phrase)*3.5
				phrase_width = round(phrase_width)
				Text 15, y_pos, phrase_width, 10, sub_number & "." & sub_phrase
				Text 15+phrase_width, y_pos, 55, 10, "CAF Answer"
				DropListBox 65+phrase_width, y_pos - 5, 35, 45, question_answers, addtl_question
				y_pos = y_pos + 20
			End If


			Text 15, y_pos, 25, 10, "Write-in:"
			If verif_status = "" Then
				EditBox 40, y_pos - 5, 435, 15, question_notes
			Else
				EditBox 40, y_pos - 5, 315, 15, question_notes
				Text 360, y_pos, 110, 10, "Q" & number & " - Verification - " & verif_status
			End If
			y_pos = y_pos + 20

			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", verif_btn
			y_pos = y_pos + 25

		ElseIf info_type = "two-part" Then
			GroupBox 5, y_pos, 475, dialog_height-5, number & "." & dialog_phrasing
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_yn
			Text 95, y_pos, 25, 10, "write-in:"
			If verif_status = "" Then
				EditBox 120, y_pos - 5, 355, 15, question_notes
			Else
				EditBox 120, y_pos - 5, 235, 15, question_notes
				Text 360, y_pos, 110, 10, "Q" & number & " - Verification - " & verif_status
			End If
			y_pos = y_pos + 20

			phrase_width = len(sub_number & "." & sub_phrase)*3.5
			phrase_width = round(phrase_width)
			Text 15, y_pos, phrase_width, 10, sub_number & "." & sub_phrase
			Text 15+phrase_width, y_pos, 55, 10, "CAF Answer"
			DropListBox 65+phrase_width, y_pos - 5, 35, 45, question_answers, addtl_question
			y_pos = y_pos + 15
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", verif_btn
			y_pos = y_pos + 20

		ElseIf info_type = "single-detail" Then
			GroupBox 5, y_pos, 475, dialog_height-5, number & "." & dialog_phrasing
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_yn

			phrase_width = len(sub_phrase & ":")*3.75
			phrase_width = round(phrase_width)
			Text 95, y_pos, phrase_width, 10, sub_phrase & ":"
			EditBox 95+phrase_width, y_pos - 5, 180-(95+phrase_width), 15, addtl_question
			Text 180, y_pos, 25, 10, "write-in:"
			If verif_status = "" Then
				EditBox 205, y_pos - 5, 270, 15, question_notes
			Else
				EditBox 205, y_pos - 5, 150, 15, question_notes
				Text 360, y_pos, 105, 10, "Q" & number & " - Verification - " & verif_status
			End If
			y_pos = y_pos + 20
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", verif_btn
			y_pos = y_pos + 20
		End If
	end sub

	public sub store_dialog_entry(question_yn, question_notes, question_interview_notes, addtl_question, TEMP_ARRAY)
		' MsgBox "Answer is Array: " & answer_is_array & vbCr & "Info Type: " & info_type

		entirely_blank = True
		If answer_is_array = false Then
			caf_answer = question_yn
			If caf_answer <> "" Then entirely_blank = false
		End If
		If answer_is_array = true Then
			If info_type <> "unea" Then
				For i = 0 to UBound(TEMP_ARRAY)
					' MsgBox "array item: " & TEMP_ARRAY(i)
					' MsgBox "number - " & number & vbCr & "TYPE - " & TypeName(item_ans_list) & vbCr & "i - " & i
					item_ans_list(i) = TEMP_ARRAY(i)
					If item_ans_list(i) <> "" Then entirely_blank = false
				Next
			End If
			If info_type = "unea" Then
				item_ans_list(0) = 	unea_1_yn
				item_detail_list(0) = trim(unea_1_amt)
				item_ans_list(1) = unea_2_yn
				item_detail_list(1) = trim(unea_2_amt)
				item_ans_list(2) = unea_3_yn
				item_detail_list(2) = trim(unea_3_amt)
				item_ans_list(3) = unea_4_yn
				item_detail_list(3) = trim(unea_4_amt)
				item_ans_list(4) = unea_5_yn
				item_detail_list(4) = trim(unea_5_amt)
				item_ans_list(5) = unea_6_yn
				item_detail_list(5) = trim(unea_6_amt)
				item_ans_list(6) = unea_7_yn
				item_detail_list(6) = trim(unea_7_amt)
				item_ans_list(7) = unea_8_yn
				item_detail_list(7) = trim(unea_8_amt)
				item_ans_list(8) = unea_9_yn
				item_detail_list(8) = trim(unea_9_amt)

				For i = 0 to UBound(item_ans_list)
					If item_ans_list(i) <> "" Then entirely_blank = false
					If item_detail_list(i) <> "" Then entirely_blank = false
				Next
			End If
		End If
		write_in_info = trim(question_notes)
		interview_notes = trim(question_interview_notes)
		sub_answer = trim(addtl_question)

		If write_in_info <> "" Then entirely_blank = false
		If interview_notes <> "" Then entirely_blank = false
		If sub_answer <> "" Then entirely_blank = false
	end sub

	public sub capture_array_verif_detail(item_index)
		'funcitonality here

		BeginDialog Dialog1, 0, 0, 396, 105, "Add Verification"
			DropListBox 60, 45, 75, 45, "Not Needed"+chr(9)+"Requested"+chr(9)+"On File"+chr(9)+"Verbal Attestation", detail_verif_status(item_index)
			EditBox 60, 65, 330, 15, detail_verif_notes(item_index)
			ButtonGroup ButtonPressed
				PushButton 340, 85, 50, 15, "Return", return_btn
				PushButton 145, 45, 50, 10, "CLEAR", clear_btn
			Text 10, 10, 380, 20, number & "." & dialog_phrasing
			If detail_source = "jobs" Then Text 10, 20, 380, 20, "Employer: " & detail_business(item_index) & "  - Employee: " & detail_resident_name(item_index) & "   - Gross Monthly Earnings: $ " & detail_monthly_amount(item_index)
			If detail_source = "assets" Then Text 10, 20, 380, 20, "Owner: " & detail_resident_name(item_index) & "  - Type: " & detail_type(item_index) & "  - Value: $ " & detail_value(item_index) & "  - Account Info: " & detail_explain(item_index)
			If detail_source = "unea" Then Text 10, 20, 380, 20, "Name: " & detail_resident_name(item_index) & "  - Type: " & detail_type(item_index) & "   - Start Date: $ " & detail_date(item_index) & "   - Amount: $ " & detail_amount(item_index) & "   - Freq.: " & detail_frequency(item_index)
			If detail_source = "shel-hest" Then Text 10, 20, 380, 20, "Type: " & detail_type(item_index) & "  - Amount: $ " & detail_amount(item_index) & "  - Frequency: " & detail_frequency(item_index)
			If detail_source = "expense" Then Text 10, 20, 380, 20, "Payer: " & detail_resident_name(item_index) & "  - Amount: $ " & detail_amount(item_index) & "  - Currently Paying: " & detail_current(item_index)
			If detail_source = "changes" Then Text 10, 20, 380, 20, "Who: " & detail_resident_name(item_index) & "  - Date: " & detail_date(item_index) & "  - Explain: " & detail_explain(item_index)
			Text 10, 50, 45, 10, "Verification: "
			Text 20, 70, 30, 10, "Details:"
		EndDialog

		Do
			dialog Dialog1
			If ButtonPressed = -1 Then ButtonPressed = return_btn
			If ButtonPressed = clear_btn Then
				verif_status = "Not Needed"
				verif_notes = ""
			End If
		Loop until ButtonPressed = return_btn
	end sub

	public sub capture_verif_detail()
		'funcitonality here

		BeginDialog Dialog1, 0, 0, 396, 95, "Add Verification"
		DropListBox 60, 35, 75, 45, "Not Needed"+chr(9)+"Requested"+chr(9)+"On File"+chr(9)+"Verbal Attestation", verif_status
		EditBox 60, 55, 330, 15, verif_notes
		ButtonGroup ButtonPressed
			PushButton 340, 75, 50, 15, "Return", return_btn
			PushButton 145, 35, 50, 10, "CLEAR", clear_btn
		Text 10, 10, 380, 20, number & "." & dialog_phrasing
		Text 10, 40, 45, 10, "Verification: "
		Text 20, 60, 30, 10, "Details:"
		EndDialog

		Do
			dialog Dialog1
			If ButtonPressed = -1 Then ButtonPressed = return_btn
			If ButtonPressed = clear_btn Then
				verif_status = "Not Needed"
				verif_notes = ""
			End If
		Loop until ButtonPressed = return_btn
	end sub

	public sub enter_case_note()
		If info_type = "standard" or info_type = "two-part" or info_type = "single-detail" or detail_array_exists = true Then
			'funcitonality here
			If caf_answer <> "" OR trim(write_in_info) <> "" OR verif_status <> "" OR trim(interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE(note_phrasing)
			q_input = "    CAF Answer - " & caf_answer
			If info_type = "single-detail" Then
				If trim(sub_answer) <> "" Then q_input = q_input & " " & sub_note_phrase & ": " & sub_answer
			End If
			If caf_answer <> "" OR trim(write_in_info) <> "" Then q_input = q_input & " (Confirmed)"
			If q_input <> "    CAF Answer - " Then CALL write_variable_in_CASE_NOTE(q_input)
			If info_type = "two-part" Then
				q_sub_verbiage = "    Q" & number & sub_number & "." & sub_note_phrase
				If sub_answer <> "" Then
					Call write_variable_in_CASE_NOTE(q_sub_verbiage)
					Call write_variable_in_CASE_NOTE("        CAF Answer - " & sub_answer)
				End If
			End If
			If trim(write_in_info) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer - " & write_in_info)
			If verif_status <> "" Then
				If trim(verif_notes) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & verif_status)
				If trim(verif_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & verif_status & ": " & verif_notes)
			End If
			If trim(interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & interview_notes)

			If detail_array_exists = true Then
				If detail_source = "shel-hest" Then
					If housing_payment <> "" Then CALL write_variable_in_CASE_NOTE("    - Housing Payment: $ " & housing_payment)
					If housing_payment = "" Then CALL write_variable_in_CASE_NOTE("    - Housing Payment: BLANK")

					utility_string = ""
					If heat_air_checkbox = unchecked Then utility_string = utility_string & "Heat/AC: No"
					If heat_air_checkbox = checked Then utility_string = utility_string & "Heat/AC: Yes"

					If electric_checkbox = unchecked Then utility_string = utility_string & "     Electric: No"
					If electric_checkbox = checked Then utility_string = utility_string & "     Electric: Yes"

					If phone_checkbox = unchecked Then utility_string = utility_string & "     Phone: No"
					If phone_checkbox = checked Then utility_string = utility_string & "     Phone: Yes"
					CALL write_variable_in_CASE_NOTE("      " & utility_string)

					If subsidy_yn <> "" or subsidy_amount <> "" Then
						If subsidy_yn <> "No" Then CALL write_variable_in_CASE_NOTE("    - Subsidy: " & subsidy_yn & "    Subsidy Amount: $ " & subsidy_amount & " /month")
						If subsidy_yn = "No" Then CALL write_variable_in_CASE_NOTE("    - Subsidy: " & subsidy_yn)
					End If
					' If subsidy_yn = "" and subsidy_amount = "" Then CALL write_variable_in_CASE_NOTE("    - Subsidy: BLANK    Subsidy Amount: $ BLANK /month")
				End If

				for each_item = 0 to UBOUND(detail_interview_notes)
					info_entered = false
					If detail_source = "jobs" Then
						If detail_business(each_item) <> "" OR detail_resident_name(each_item) <> "" OR detail_monthly_amount(each_item) <> "" OR detail_hourly_wage(each_item) <> "" OR detail_hours_per_week(each_item) <> "" Then
							CALL write_variable_in_CASE_NOTE("    *Employer: " & detail_business(each_item) & " for " & detail_resident_name(each_item) & " monthly earnings $" & detail_monthly_amount(each_item))
							If detail_hourly_wage(each_item) <> "" OR detail_hours_per_week(each_item) <> "" Then CALL write_variable_in_CASE_NOTE("     Hourly Wage: $ " & detail_hourly_wage(each_item) & " hours/week: " & detail_hours_per_week(each_item))
							info_entered = true
						End If
					ElseIf detail_source = "assets" Then
						If detail_resident_name(each_item) <> "" OR detail_type(each_item) <> "" OR detail_value(each_item) <> "" Then
							If detail_explain(each_item) = "" Then CALL write_variable_in_CASE_NOTE("    *Owner: " & detail_resident_name(each_item) & " - Type: " & detail_type(each_item) & " - Value $" & detail_value(each_item))
							If detail_explain(each_item) <> "" Then CALL write_variable_in_CASE_NOTE("    *Owner: " & detail_resident_name(each_item) & " - Type: " & detail_type(each_item) & " - Value $" & detail_value(each_item) & " - Info: " & detail_explain(each_item))
							info_entered = true
						End If
					ElseIf detail_source = "unea" Then
						If detail_resident_name(each_item) <> "" OR detail_type(each_item) <> "" OR detail_date(each_item) <> "" OR detail_amount(each_item) <> "" OR detail_frequency(each_item) <> "" Then
							CALL write_variable_in_CASE_NOTE("    *Income type: " & detail_type(each_item) & " for " & detail_resident_name(each_item) & " Amount $" & detail_amount(each_item) &  " freq: " & detail_frequency(each_item))
							If detail_date(each_item) <> "" Then CALL write_variable_in_CASE_NOTE("     Start Date: " & detail_date(each_item))
							info_entered = true
						End If
					ElseIf detail_source = "shel-hest" Then
						If detail_type(each_item) <> "" OR detail_amount(each_item) <> "" OR detail_frequency(each_item) <> "" Then
							CALL write_variable_in_CASE_NOTE("    *Expense type: " & detail_type(each_item) & " amount $" & detail_amount(each_item) &  " freq: " & detail_frequency(each_item))
							info_entered = true
						End If
					ElseIf detail_source = "expense" Then
						If detail_resident_name(each_item) <> "" OR detail_amount(each_item) <> "" OR detail_current(each_item) <> "" Then
							CALL write_variable_in_CASE_NOTE("    *Payer: " & detail_resident_name(each_item) & " amount $" & detail_amount(each_item) &  ". Paying? " & detail_current(each_item))
							info_entered = true
						End If
					ElseIf detail_source = "winnings" Then
						If detail_resident_name(each_item) <> "" or detail_amount(each_item) <> "" or detail_date(each_item) <> "" Then
							CALL write_variable_in_CASE_NOTE("    *Winner: " & detail_resident_name(each_item) & " amount $" & detail_amount(each_item) &  ". Date of Win: " & detail_date(each_item))
							info_entered = true
						End If
					ElseIf detail_source = "changes" Then
						If detail_resident_name(each_item) <> "" OR detail_date(each_item) <> "" OR detail_explain(each_item) <> "" Then
							CALL write_variable_in_CASE_NOTE("    *Date: " & detail_date(each_item) & " change: " & detail_explain(each_item) & " person: " & detail_resident_name(each_item))
							info_entered = true
						End If
					End If
					If info_entered = true Then
						If detail_verif_status(each_item) <> "" Then
							If trim(detail_verif_notes(each_item)) = "" Then CALL write_variable_in_CASE_NOTE("     Verification: " & detail_verif_status(each_item))
							If trim(detail_verif_notes(each_item)) <> "" Then CALL write_variable_in_CASE_NOTE("     Verification: " & detail_verif_status(each_item) & ": " & detail_verif_notes(each_item))
						End If
						If trim(detail_write_in_info(each_item)) <> "" Then CALL write_variable_in_CASE_NOTE("     WriteIn Answer: " & detail_write_in_info(each_item))
						If trim(detail_interview_notes(each_item)) <> "" Then CALL write_variable_in_CASE_NOTE("     INTVW NOTES: " & detail_interview_notes(each_item))
					End If
				next
				If detail_source = "shel-hest" Then

				End If
			End If
		' ElseIf info_type = "two-part" Then
		' ElseIf info_type = "jobs" Then
		' ElseIf info_type = "single-detail" Then
		ElseIf info_type = "unea" Then
			If entirely_blank = false Then
				Call write_variable_in_CASE_NOTE(note_phrasing)
				CALL write_variable_in_CASE_NOTE("    CAF Answer:")

				for i = 0 to 8
					item_ans_list(i) = left(item_ans_list(i) & "   ", 5)
					If trim(item_detail_list(i)) <> "" Then item_detail_list(i) = left("$" & item_detail_list(i) & ".00       ", 8)
				next
				CALL write_variable_in_CASE_NOTE("    RSDI - " & item_ans_list(0) & " " & item_detail_list(0) & "   UI - " & item_ans_list(1) & " " & item_detail_list(1) & " Tribal - " & item_ans_list(2) & " " & item_detail_list(2))
				CALL write_variable_in_CASE_NOTE("     SSI - " & item_ans_list(3) & " " & item_detail_list(3) & "   WC - " & item_ans_list(4) & " " & item_detail_list(4) & "   CSES - " & item_ans_list(5) & " " & item_detail_list(5))
				CALL write_variable_in_CASE_NOTE("      VA - " & item_ans_list(6) & " " & item_detail_list(6) & "  Ret - " & item_ans_list(7) & " " & item_detail_list(7) & "  Other - " & item_ans_list(8) & " " & item_detail_list(8))
			End If
			If trim(write_in_info) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer - " & write_in_info)
			If verif_status <> "" Then
				If trim(verif_notes) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & verif_status)
				If trim(verif_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & verif_status & ": " & verif_notes)
			End If
			If trim(interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & interview_notes)
		ElseIf info_type = "housing" or info_type = "utilities" or info_type = "assets" or info_type = "msa" or info_type = "stwk" Then
			If entirely_blank = false Then
				Call write_variable_in_CASE_NOTE(note_phrasing)
				CALL write_variable_in_CASE_NOTE("    CAF Answer:")

				for i = 0 to UBound(item_ans_list)
					If IsNumeric(item_ans_list(i)) = True Then
						If item_ans_list(i) = 1 Then
							item_ans_list(i) = "Yes"
						ElseIf item_ans_list(i) = 0 Then
							item_ans_list(i) = "No"
						End If
					End If
					item_ans_list(i) = left(item_ans_list(i) & "   ", 5)
				next
				If info_type = "housing" Then
					spaces_1 = "       "
					spaces_2 = "                        "
					If sub_phrase = "" Then
						CALL write_variable_in_CASE_NOTE(spaces_1 & "Rent - " & item_ans_list(0) &  " Rental Subsidy - " & item_ans_list(1) & "  Mortgage - " & item_ans_list(2) & "    Taxes - " & item_ans_list(3))
						CALL write_variable_in_CASE_NOTE(spaces_2 & "Assoc Fees - " & item_ans_list(4) & "Room/Board - " & item_ans_list(5)    & "Insurance - " & item_ans_list(6))
					Else
						CALL write_variable_in_CASE_NOTE(spaces_1 & "      Rent - " & item_ans_list(0) & "  Mortgage - " & item_ans_list(1) & "     Taxes - " & item_ans_list(2))
						CALL write_variable_in_CASE_NOTE(spaces_1 & "Assoc Fees - " & item_ans_list(3) & "Room/Board - " & item_ans_list(4) & " Insurance - " & item_ans_list(5))
					End If
				End If
				If info_type = "utilities" Then
					CALL write_variable_in_CASE_NOTE("        Heat/AC - " & item_ans_list(0) & " Electric - " & item_ans_list(1) & " Cooking Fuel - " & item_ans_list(2))
					CALL write_variable_in_CASE_NOTE("    Water/Sewer - " & item_ans_list(3) & "  Garbage - " & item_ans_list(4) & "        Phone - " & item_ans_list(5))
					If sub_phrase = "" Then CALL write_variable_in_CASE_NOTE("    LIHEAP/Energy Assistance in past 12 months - " & item_ans_list(6))
				End If
				If info_type = "assets" Then
					CALL write_variable_in_CASE_NOTE("      Cash - " & item_ans_list(0) & " Bank Accounts - " & item_ans_list(1))
					CALL write_variable_in_CASE_NOTE("    Stocks - " & item_ans_list(2) & "      Vehicles - " & item_ans_list(3))
					CALL write_variable_in_CASE_NOTE("         Electronic Payment Card - " & item_ans_list(4))
				End If
				If info_type = "msa" or info_type = "stwk" Then
					If info_type = "msa" Then
						CALL write_variable_in_CASE_NOTE("    REP Payee Fees - " & item_ans_list(0) & "           Guard Fees - " & item_ans_list(1))
						CALL write_variable_in_CASE_NOTE("      Special Diet - " & item_ans_list(2) & "   High Housing Costs - " & item_ans_list(3))
					End If
					If info_type = "stwk" Then
						CALL write_variable_in_CASE_NOTE("           Stop Working - " & item_ans_list(0) & "     Refuse a Job - " & item_ans_list(1))
						CALL write_variable_in_CASE_NOTE("    Request Fewer Hours - " & item_ans_list(2) & "           Strike - " & item_ans_list(3))
					End If

				End If
			End If
			If sub_phrase <> "" Then
				q_sub_verbiage = "    Q" & number & sub_number & "." & sub_note_phrase
				If sub_answer <> "" Then
					Call write_variable_in_CASE_NOTE(q_sub_verbiage)
					Call write_variable_in_CASE_NOTE("        CAF Answer - " & sub_answer)
				End If
			End If
			If trim(write_in_info) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer - " & write_in_info)
			If verif_status <> "" Then
				If trim(verif_notes) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & verif_status)
				If trim(verif_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & verif_status & ": " & verif_notes)
			End If
			If trim(interview_notes) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & interview_notes)
		End If
	end sub

	public sub add_to_wif()
		'funcitonality here
		objSelection.TypeText doc_phrasing & vbCr

		If info_type = "standard" or info_type = "two-part" or info_type = "single-detail" or detail_array_exists = true Then
			objSelection.TypeText chr(9) & "CAF Answer: " & caf_answer & vbCr
			If info_type = "two-part" Then
				objSelection.TypeText "Q "& number&"."&sub_number&". "&sub_phrase & vbCr
				objSelection.TypeText chr(9) & "CAF Answer: " & sub_answer & vbCr
			End If
			If info_type = "single-detail" Then
				If sub_answer <> "" Then objSelection.TypeText chr(9) & sub_note_phrase & ": " & sub_answer & vbCr
				If sub_answer = "" Then objSelection.TypeText chr(9) & sub_note_phrase & ": NONE LISTED" & vbCr
			End If

			If detail_array_exists = true Then
				detail_added = False
				If detail_source = "jobs" Then
					for each_item = 0 to UBound(detail_interview_notes)
						none_added_info = "THERE ARE NO JOBS ENTERED."
						If detail_business(each_item) <> "" OR detail_resident_name(each_item) <> "" OR detail_monthly_amount(each_item) <> "" OR detail_hourly_wage(each_item) <> "" Then
							detail_added = true

							all_the_tables = UBound(TABLE_ARRAY) + 1
							ReDim Preserve TABLE_ARRAY(all_the_tables)
							Set objRange = objSelection.Range					'range is needed to create tables
							objDoc.Tables.Add objRange, 8, 1					'This sets the rows and columns needed row then column'
							set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
							table_count = table_count + 1

							TABLE_ARRAY(array_counters).AutoFormat(16)							'This adds the borders to the table and formats it
							TABLE_ARRAY(array_counters).Columns(1).Width = 400

							for row = 1 to 7 Step 2
								TABLE_ARRAY(array_counters).Cell(row, 1).SetHeight 10, 2
							Next
							for row = 2 to 8 Step 2
								TABLE_ARRAY(array_counters).Cell(row, 1).SetHeight 15, 2
							Next

							For row = 1 to 2
								TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 3, TRUE

								TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 200, 2
								TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 90, 2
								TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 110, 2
							Next
							For col = 1 to 3
								TABLE_ARRAY(array_counters).Cell(1, col).Range.Font.Size = 6
								TABLE_ARRAY(array_counters).Cell(2, col).Range.Font.Size = 12
							Next
							TABLE_ARRAY(array_counters).Cell(3, 1).Range.Font.Size = 6
							TABLE_ARRAY(array_counters).Cell(4, 1).Range.Font.Size = 12
							TABLE_ARRAY(array_counters).Cell(5, 1).Range.Font.Size = 6
							TABLE_ARRAY(array_counters).Cell(6, 1).Range.Font.Size = 12
							TABLE_ARRAY(array_counters).Cell(7, 1).Range.Font.Size = 6
							TABLE_ARRAY(array_counters).Cell(8, 1).Range.Font.Size = 12

							TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = "EMPLOYEE NAME"
							TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "HOURLY WAGE"
							TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = "GROSS MONTHLY EARNINGS"
							TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text = detail_resident_name(each_item)
							TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text = detail_hourly_wage(each_item)
							TABLE_ARRAY(array_counters).Cell(2, 3).Range.Text = detail_monthly_amount(each_item)

							TABLE_ARRAY(array_counters).Cell(3, 1).Range.Text = "EMPLOYER/BUSINESS NAME"
							TABLE_ARRAY(array_counters).Cell(4, 1).Range.Text = detail_business(each_item)

							TABLE_ARRAY(array_counters).Cell(5, 1).Range.Text = "CAF NOTES"
							TABLE_ARRAY(array_counters).Cell(6, 1).Range.Text = detail_write_in_info(each_item)

							TABLE_ARRAY(array_counters).Cell(7, 1).Range.Text = "INTERVIEW NOTES"
							TABLE_ARRAY(array_counters).Cell(8, 1).Range.Text = detail_interview_notes(each_item)

							objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
							' objSelection.TypeParagraph()						'adds a line between the table and the next information

							array_counters = array_counters + 1

							objSelection.TypeText "Verification: " & detail_verif_status(each_item) & " - " & detail_verif_notes(each_item) & vbCR
						End If
					next
				End If


				If detail_source = "assets" OR  detail_source = "unea" OR detail_source = "shel-hest" OR detail_source = "expense" OR detail_source = "winnings" OR detail_source = "changes" Then
					If detail_source = "assets" Then none_added_info = "THERE ARE NO ASSETS ENTERED"
					If detail_source = "unea" Then none_added_info = "THERE IS NO UNEARNED INCOME ENTERED"
					If detail_source = "shel-hest" Then none_added_info = "THERE IS NO HOUSING OR UTILITIES ENTERED"
					If detail_source = "expense" Then none_added_info = "THERE ARE NO EXPENSES ENTERED"
					If detail_source = "changes" Then none_added_info = "THERE ARE NO CHANGES ENTERED"
					row_count = 1
					for each_item = 0 to UBound(detail_interview_notes)
						If detail_source = "assets" Then
							If detail_resident_name(each_item) <> "" OR detail_type(each_item) <> "" OR detail_value(each_item) <> "" or detail_explain(each_item) <> ""  Then row_count = row_count + 2
						End If
						If detail_source = "unea" Then
							If detail_resident_name(each_item) <> "" OR detail_type(each_item) <> "" OR detail_date(each_item) <> "" OR detail_amount(each_item) <> "" OR detail_frequency(each_item) <> "" Then row_count = row_count + 2
						End If
						If detail_source = "shel-hest" Then
							If detail_type(each_item) <> "" OR  detail_amount(each_item) <> "" OR detail_frequency(each_item) <> "" Then row_count = row_count + 2
						End If
						If detail_source = "expense" Then
							If detail_resident_name(each_item) <> "" OR detail_amount(each_item) <> "" OR detail_current(each_item) <> "" Then row_count = row_count + 2
						End If
						If detail_source = "winnings" Then
							If detail_resident_name(each_item) <> "" or detail_amount(each_item) <> "" or detail_date(each_item) <> "" Then row_count = row_count + 2
						End If
						If detail_source = "changes" Then
							If detail_resident_name(each_item) <> "" OR detail_date(each_item) <> "" OR detail_explain(each_item) <> "" Then row_count = row_count + 5
						End If
					next
					If detail_source = "changes" and row_count > 1 Then row_count = row_count - 1

					If row_count <> 1 Then
						detail_added = True
						all_the_tables = UBound(TABLE_ARRAY) + 1
						ReDim Preserve TABLE_ARRAY(all_the_tables)
						Set objRange = objSelection.Range					'range is needed to create tables
						objDoc.Tables.Add objRange, row_count, 1					'This sets the rows and columns needed row then column'
						set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
						table_count = table_count + 1

						TABLE_ARRAY(array_counters).AutoFormat(16)							'This adds the borders to the table and formats it
						TABLE_ARRAY(array_counters).Columns(1).Width = 500
						col_count = 0

						If detail_source <> "changes" Then
							row = 1
							Do
								If row <> 1 Then
									TABLE_ARRAY(array_counters).Cell(row, 1).SetHeight 15, 2
									TABLE_ARRAY(array_counters).Cell(row+1, 1).SetHeight 30, 2
								Else
									TABLE_ARRAY(array_counters).Cell(row, 1).SetHeight 10, 2
								End If
								If detail_source = "assets" or detail_source = "shel-hest" or detail_source = "expense" or detail_source = "winnings" Then
									TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 3, TRUE

									TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 200, 2
									TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 190, 2
									TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 110, 2
									col_count = 3
								ElseIf detail_source = "unea" Then
									TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 5, TRUE

									TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 125, 2
									TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 125, 2
									TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 75, 2
									TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 75, 2
									TABLE_ARRAY(array_counters).Cell(row, 5).SetWidth 100, 2
									col_count = 5

								End If
								If row <> 1 Then
									For col = 1 to col_count
										TABLE_ARRAY(array_counters).Cell(row, col).Range.Font.Size = 12
									Next
									TABLE_ARRAY(array_counters).Cell(row+1, 1).Range.Font.Size = 12
								Else
									For col = 1 to col_count
										TABLE_ARRAY(array_counters).Cell(1, col).Range.Font.Size = 6
									Next
								End If
								If row = 1 Then
									row = 2
								Else
									row = row + 2
								End If
							Loop until row >= row_count
						End If

						If detail_source = "changes" Then
							row = 1
							Do
								TABLE_ARRAY(array_counters).Cell(row, 1).SetHeight 10, 2
								TABLE_ARRAY(array_counters).Cell(row+1, 1).SetHeight 15, 2
								TABLE_ARRAY(array_counters).Cell(row+2, 1).SetHeight 10, 2
								TABLE_ARRAY(array_counters).Cell(row+3, 1).SetHeight 15, 2
								TABLE_ARRAY(array_counters).Cell(row+4, 1).SetHeight 30, 2

								TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 2, TRUE
								TABLE_ARRAY(array_counters).Rows(row+1).Cells.Split 1, 2, TRUE

								TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 425, 2
								TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 75, 2

								TABLE_ARRAY(array_counters).Cell(row+1, 1).SetWidth 425, 2
								TABLE_ARRAY(array_counters).Cell(row+1, 2).SetWidth 75, 2

								TABLE_ARRAY(array_counters).Cell(row, 1).Range.Font.Size = 6
								TABLE_ARRAY(array_counters).Cell(row, 2).Range.Font.Size = 6
								TABLE_ARRAY(array_counters).Cell(row+1, 1).Range.Font.Size = 12
								TABLE_ARRAY(array_counters).Cell(row+1, 2).Range.Font.Size = 12
								TABLE_ARRAY(array_counters).Cell(row+2, 1).Range.Font.Size = 6
								TABLE_ARRAY(array_counters).Cell(row+3, 1).Range.Font.Size = 12
								TABLE_ARRAY(array_counters).Cell(row+4, 1).Range.Font.Size = 12

								row = row + 5
							Loop until row >= row_count

							row = 1
							for each_item = 0 to UBound(detail_interview_notes)
								If detail_resident_name(each_item) <> "" OR detail_date(each_item) <> "" OR detail_explain(each_item) <> "" Then
									TABLE_ARRAY(array_counters).Cell(row, 1).Range.Text = "Who?"
									TABLE_ARRAY(array_counters).Cell(row, 2).Range.Text = "Date of Change"

									TABLE_ARRAY(array_counters).Cell(row+1, 1).Range.Text = detail_resident_name(each_item)
									TABLE_ARRAY(array_counters).Cell(row+1, 2).Range.Text = detail_date(each_item)

									TABLE_ARRAY(array_counters).Cell(row+2, 1).Range.Text = "Explain the change"
									TABLE_ARRAY(array_counters).Cell(row+3, 1).Range.Text = detail_explain(each_item)

									notes_detail = "Write-In: " & detail_write_in_info(each_item) & vbCr
									notes_detail = notes_detail & "Notes: " & detail_interview_notes(each_item)
									If detail_verif_status(each_item) <> "" or detail_verif_notes(each_item) <> "" Then
										notes_detail = notes_detail & vbCr & "Verification: " & detail_verif_status(each_item) & " - " & detail_verif_notes(each_item)
										TABLE_ARRAY(array_counters).Cell(row+1, 1).SetHeight 45, 2
									End If
									TABLE_ARRAY(array_counters).Cell(row+4, 1).Range.ParagraphFormat.SpaceAfter = 0
									TABLE_ARRAY(array_counters).Cell(row+4, 1).Range.Text = notes_detail
									row = row + 5

								End If
							next

						End If

						If detail_source = "assets" Then
							TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = "Owner(s) name"
							TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "Information"
							TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = "Value or amount in account"
						ElseIf detail_source = "unea" Then
							TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = "Name"
							TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "Type of Income"
							TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = "Start date"
							TABLE_ARRAY(array_counters).Cell(1, 4).Range.Text = "Amount"
							TABLE_ARRAY(array_counters).Cell(1, 5).Range.Text = "How often received"
						ElseIf detail_source = "shel-hest" Then
							TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = "Expense"
							TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "Amount"
							TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = "How often?"
						ElseIf detail_source = "winnings" Then
							TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = "Name of Winner"
							TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "Amount"
							TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = "Date of Win"
						ElseIf detail_source = "expense" Then
							TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = "Name of person paying"
							TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "Monthly Amount"
							TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = "Currently paying"
						End If


						row = 2
						If detail_source = "assets" or detail_source = "shel-hest" or detail_source = "expense" or detail_source = "unea" or detail_source = "winnings" Then
							for each_item = 0 to UBound(detail_interview_notes)
								enter_note = false
								If detail_source = "assets" Then
									If detail_resident_name(each_item) <> "" OR detail_type(each_item) <> "" OR detail_value(each_item) <> "" Then
										TABLE_ARRAY(array_counters).Cell(row, 1).Range.Text = detail_resident_name(each_item)
										TABLE_ARRAY(array_counters).Cell(row, 2).Range.Text = detail_type(each_item) & " - Info: " & detail_explain(each_item)
										TABLE_ARRAY(array_counters).Cell(row, 3).Range.Text = "$ " & detail_value(each_item)
										enter_note = true
									End If
								End If
								If detail_source = "unea" Then
									If detail_resident_name(each_item) <> "" OR detail_type(each_item) <> "" OR detail_date(each_item) <> "" OR detail_amount(each_item) <> "" OR detail_frequency(each_item) <> "" Then
										TABLE_ARRAY(array_counters).Cell(row, 1).Range.Text = detail_resident_name(each_item)
										TABLE_ARRAY(array_counters).Cell(row, 2).Range.Text = detail_type(each_item)
										TABLE_ARRAY(array_counters).Cell(row, 3).Range.Text = detail_date(each_item)
										TABLE_ARRAY(array_counters).Cell(row, 4).Range.Text = "$ " & detail_amount(each_item)
										TABLE_ARRAY(array_counters).Cell(row, 5).Range.Text = detail_frequency(each_item)
										enter_note = true
									End If
								End If
								If detail_source = "shel-hest" Then
									If detail_type(each_item) <> "" OR  detail_amount(each_item) <> "" OR detail_frequency(each_item) <> "" Then
										TABLE_ARRAY(array_counters).Cell(row, 1).Range.Text = detail_type(each_item)
										TABLE_ARRAY(array_counters).Cell(row, 2).Range.Text = "$ " & detail_amount(each_item)
										TABLE_ARRAY(array_counters).Cell(row, 3).Range.Text = detail_frequency(each_item)
										enter_note = true
									End If
								End If
								If detail_source = "expense" Then
									If detail_resident_name(each_item) <> "" OR detail_amount(each_item) <> "" OR detail_current(each_item) <> "" Then
										TABLE_ARRAY(array_counters).Cell(row, 1).Range.Text = detail_resident_name(each_item)
										TABLE_ARRAY(array_counters).Cell(row, 2).Range.Text = "$ " & detail_amount(each_item)
										TABLE_ARRAY(array_counters).Cell(row, 3).Range.Text = detail_current(each_item)
										enter_note = true
									End If
								End If
								If detail_source = "winnings" Then
									If detail_resident_name(each_item) <> "" or detail_amount(each_item) <> "" or detail_date(each_item) <> "" Then
										TABLE_ARRAY(array_counters).Cell(row, 1).Range.Text = detail_resident_name(each_item)
										TABLE_ARRAY(array_counters).Cell(row, 2).Range.Text = "$ " & detail_amount(each_item)
										TABLE_ARRAY(array_counters).Cell(row, 3).Range.Text = detail_date(each_item)
										enter_note = true
									End If
								End If
								If enter_note = true Then
									notes_detail = "Write-In: " & detail_write_in_info(each_item) & vbCr
									notes_detail = notes_detail & "Notes: " & detail_interview_notes(each_item)
									If detail_verif_status(each_item) <> "" or detail_verif_notes(each_item) <> "" Then
										notes_detail = notes_detail & vbCr & "Verification: " & detail_verif_status(each_item) & " - " & detail_verif_notes(each_item)
										TABLE_ARRAY(array_counters).Cell(row+1, 1).SetHeight 45, 2
									End If
									TABLE_ARRAY(array_counters).Cell(row+1, 1).Range.ParagraphFormat.SpaceAfter = 0
									TABLE_ARRAY(array_counters).Cell(row+1, 1).Range.Text = notes_detail
									row = row + 2
								End If
							next
						End If
						objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
						array_counters = array_counters + 1
					End If
				End If
				If detail_added = FALSE Then objSelection.TypeText chr(9) & none_added_info & vbCr

			End If

			If detail_source = "shel-hest" Then
				If housing_payment <> "" Then objSelection.TypeText chr(9) & "Housing Payment: $ " & housing_payment & vbCr
				If housing_payment = "" Then objSelection.TypeText chr(9) & "Housing Payment: BLANK" & vbCr

				objSelection.Font.Name = "Wingdings"
				If heat_air_checkbox = unchecked Then objSelection.TypeText chr(9) & chr(111)
				If heat_air_checkbox = checked Then objSelection.TypeText chr(9) & chr(120)
				objSelection.Font.Name = "Arial"
				objSelection.TypeText " Heat/Air Conditioning"

				objSelection.Font.Name = "Wingdings"
				If electric_checkbox = unchecked Then objSelection.TypeText chr(9) & chr(111)
				If electric_checkbox = checked Then objSelection.TypeText chr(9) & chr(120)
				objSelection.Font.Name = "Arial"
				objSelection.TypeText " Electricity"

				objSelection.Font.Name = "Wingdings"
				If phone_checkbox = unchecked Then objSelection.TypeText chr(9) & chr(111)
				If phone_checkbox = checked Then objSelection.TypeText chr(9) & chr(120)
				objSelection.Font.Name = "Arial"
				objSelection.TypeText " Telephone" & vbCr

				If subsidy_yn <> "" or subsidy_amount <> "" Then
					If subsidy_yn <> "No" Then objSelection.TypeText chr(9) & "Subsidy: " & subsidy_yn & "    Subsidy Amount: $ " & subsidy_amount & " /month" & vbCr
					If subsidy_yn = "No" Then objSelection.TypeText chr(9) & "Subsidy: " & subsidy_yn & vbCr
				End If
				If subsidy_yn = "" and subsidy_amount = "" Then objSelection.TypeText chr(9) & "Subsidy: BLANK    Subsidy Amount: $ BLANK /month" & vbCr
			End If

			If write_in_info <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & write_in_info & vbCr
			If verif_status <> "Mot Needed" AND verif_status <> "" Then objSelection.TypeText chr(9) & "Verification: " & verif_status & vbCr
			If verif_notes <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & verif_notes & vbCr
			If caf_answer <> "" OR trim(write_in_info) <> "" Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
			If interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & interview_notes & vbCR
		ElseIf info_type = "unea" or info_type = "housing" or info_type = "utilities" or info_type = "assets" or info_type = "msa" or info_type = "stwk" Then

			all_the_tables = UBound(TABLE_ARRAY) + 1
			ReDim Preserve TABLE_ARRAY(all_the_tables)
			' MsgBox "all_the_tables - " & all_the_tables & vbCr & "array_counters - " & array_counters
			Set objRange = objSelection.Range					'range is needed to create tables
			If info_type = "unea" Then objDoc.Tables.Add objRange, 5, 1					'This sets the rows and columns needed row then column'
			If info_type = "housing" Then objDoc.Tables.Add objRange, 4, 1					'This sets the rows and columns needed row then column'
			If info_type = "utilities" Then objDoc.Tables.Add objRange, 3, 1					'This sets the rows and columns needed row then column'
			If info_type = "msa" or info_type = "stwk" Then objDoc.Tables.Add objRange, 2, 1					'This sets the rows and columns needed row then column'
			If info_type = "assets" Then objDoc.Tables.Add objRange, 3, 1					'This sets the rows and columns needed row then column'
			set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
			table_count = table_count + 1

			'note that this table does not use an autoformat - which is why there are no borders on this table.'

			If info_type = "unea" Then
				TABLE_ARRAY(array_counters).Columns(1).Width = 500
				numb_of_rows = 5
				number_of_columns = 6
				For row = 1 to 4
					TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 6, TRUE

					TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 50, 2
					TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 125, 2
					TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 75, 2
					TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 50, 2
					TABLE_ARRAY(array_counters).Cell(row, 5).SetWidth 125, 2
					TABLE_ARRAY(array_counters).Cell(row, 6).SetWidth 75, 2
				Next
				TABLE_ARRAY(array_counters).Rows(5).Cells.Split 1, 3, TRUE

				TABLE_ARRAY(array_counters).Cell(5, 1).SetWidth 50, 2
				TABLE_ARRAY(array_counters).Cell(5, 2).SetWidth 175, 2
				TABLE_ARRAY(array_counters).Cell(5, 3).SetWidth 75, 2

			ElseIf info_type = "housing" Then
				numb_of_rows = 4
				number_of_columns = 4				'note that this table does not use an autoformat - which is why there are no borders on this table.'
				TABLE_ARRAY(array_counters).Columns(1).Width = 520

				For row = 1 to 3
					TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 4, TRUE

					TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 90, 2
					TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 170, 2
					TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 90, 2
					TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 170, 2
				Next
				TABLE_ARRAY(array_counters).Rows(4).Cells.Split 1, 2, TRUE

				TABLE_ARRAY(array_counters).Cell(4, 1).SetWidth 90, 2
				TABLE_ARRAY(array_counters).Cell(4, 2).SetWidth 430, 2
			ElseIf info_type = "utilities" Then
				numb_of_rows = 3
				number_of_columns = 6
				'note that this table does not use an autoformat - which is why there are no borders on this table.'
				TABLE_ARRAY(array_counters).Columns(1).Width = 525

				For row = 1 to 2
					TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 6, TRUE

					TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 75, 2
					TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 100, 2
					TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 75, 2
					TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 100, 2
					TABLE_ARRAY(array_counters).Cell(row, 5).SetWidth 75, 2
					TABLE_ARRAY(array_counters).Cell(row, 6).SetWidth 100, 2
				Next
				TABLE_ARRAY(array_counters).Rows(3).Cells.Split 1, 2, TRUE

				TABLE_ARRAY(array_counters).Cell(3, 1).SetWidth 75, 2
				TABLE_ARRAY(array_counters).Cell(3, 2).SetWidth 450, 2

			ElseIf info_type = "assets" Then
				numb_of_rows = 3
				number_of_columns = 4

				'note that this table does not use an autoformat - which is why there are no borders on this table.'
				TABLE_ARRAY(array_counters).Columns(1).Width = 520

				For row = 1 to 2
					TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 4, TRUE

					TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 90, 2
					TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 170, 2
					TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 90, 2
					TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 170, 2
				Next
				TABLE_ARRAY(array_counters).Rows(3).Cells.Split 1, 2, TRUE
				TABLE_ARRAY(array_counters).Cell(3, 1).SetWidth 90, 2
				TABLE_ARRAY(array_counters).Cell(3, 2).SetWidth 430, 2

			ElseIf info_type = "msa" or info_type = "stwk" Then
				numb_of_rows = 2
				number_of_columns = 4

				'note that this table does not use an autoformat - which is why there are no borders on this table.'
				TABLE_ARRAY(array_counters).Columns(1).Width = 520

				For row = 1 to 2
					TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 4, TRUE

					TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 90, 2
					TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 170, 2
					TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 90, 2
					TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 170, 2
				Next
			End If

			row = 1
			col = 1
			for i = 0 to UBound(item_ans_list)
				' MsgBox "item ans - " & item_ans_list(i) & vbCr & "i - " & i & vbCr & vbCr & "array_counters - " & array_counters & vbCr & "row - " & row & vbCr & "col - " & col
				If IsNumeric(item_ans_list(i)) = True Then
					If item_ans_list(i) = 1 Then
						TABLE_ARRAY(array_counters).Cell(row, col).Range.Text = "Yes"
					ElseIf item_ans_list(i) <> "" Then
						TABLE_ARRAY(array_counters).Cell(row, col).Range.Text = "No"
					End If
				Else
					TABLE_ARRAY(array_counters).Cell(row, col).Range.Text = item_ans_list(i)
				End If
				TABLE_ARRAY(array_counters).Cell(row, col+1).Range.Text = item_note_info_list(i)
				col = col + 2
				If info_type = "unea" Then
					TABLE_ARRAY(array_counters).Cell(row, col).Range.Text = "$ " & item_detail_list(i)
					col = col + 1
				End If
				If col > number_of_columns Then
					row = row + 1
					col = 1
				End If
				' MsgBox "Filled in?"
			next


			objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
			array_counters = array_counters + 1
			If sub_phrase <> "" Then
				objSelection.TypeText "Q "& number&"."&sub_number&". "&sub_phrase & vbCr
				objSelection.TypeText chr(9) & "CAF Answer: " & sub_answer & vbCr
			End If

			If write_in_info <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & write_in_info & vbCr
			If verif_status <> "Mot Needed" AND verif_status <> "" Then objSelection.TypeText chr(9) & "Verification: " & verif_status & vbCr
			If verif_notes <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & verif_notes & vbCr

			If entirely_blank = false Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
			If interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & interview_notes & vbCR
		End If
	end sub

	public sub restore_info(node)
		Set subNode = node.SelectSingleNode("number")
		If Not subNode Is Nothing Then
			number = subNode.Text
			number = number * 1
		End If


		Set subNode = node.SelectSingleNode("dialogPhrasing")
		If Not subNode Is Nothing Then dialog_phrasing = subNode.Text
		Set subNode = node.SelectSingleNode("notePhrasing")
		If Not subNode Is Nothing Then note_phrasing = subNode.Text
		Set subNode = node.SelectSingleNode("docPhrasing")
		If Not subNode Is Nothing Then doc_phrasing = subNode.Text
		Set subNode = node.SelectSingleNode("subNumber")
		If Not subNode Is Nothing Then sub_number = subNode.Text
		Set subNode = node.SelectSingleNode("subPhrase")
		If Not subNode Is Nothing Then sub_phrase = subNode.Text
		Set subNode = node.SelectSingleNode("subNotePhrase")
		If Not subNode Is Nothing Then sub_note_phrase = subNode.Text
		Set subNode = node.SelectSingleNode("subAnswer")
		If Not subNode Is Nothing Then sub_answer = subNode.Text
		Set subNode = node.SelectSingleNode("infoType")
		If Not subNode Is Nothing Then info_type = subNode.Text
		Set subNode = node.SelectSingleNode("cafAnswer")
		If Not subNode Is Nothing Then caf_answer = subNode.Text
		Set subNode = node.SelectSingleNode("answerIsArray")
		If Not subNode Is Nothing Then
			answer_is_array = subNode.Text
			If answer_is_array = "0" Then answer_is_array = False
			If answer_is_array = "-1" Then answer_is_array = True
		End If
		Set subNode = node.SelectSingleNode("makeArrayCheckboxes")
		If Not subNode Is Nothing Then
			make_array_checkboxes = subNode.Text
			If make_array_checkboxes = "0" Then make_array_checkboxes = False
			If make_array_checkboxes = "-1" Then make_array_checkboxes = True
		End If
		Set subNode = node.SelectSingleNode("writeInInfo")
		If Not subNode Is Nothing Then write_in_info = subNode.Text
		Set subNode = node.SelectSingleNode("interviewNotes")
		If Not subNode Is Nothing Then interview_notes = subNode.Text

		set subNodesList = node.SelectNodes("/form/question[number='"& number &"']/itemInfoList")
		temp_array = ""
		If Not subNodesList Is Nothing Then
			for each nodeItem in subNodesList
				' MsgBox "subNODE ITEM" & vbCr &  nodeItem.Text & vbCr & vbCr & "number - " & number
				temp_array = temp_array & nodeItem.Text & "~!~"
			next
		End If
		If temp_array <> "" Then
			temp_array = left(temp_array, len(temp_array)-3)
			item_info_list = split(temp_array, "~!~")
		End If
		set subNodesList = nothing



		set subNodesList = node.SelectNodes("/form/question[number='"& number &"']/itemNoteInfoList")
		temp_array = ""
		If Not subNodesList Is Nothing Then
			for each nodeItem in subNodesList
				temp_array = temp_array & nodeItem.Text & "~!~"
			next
		End If
		If temp_array <> "" Then
			temp_array = left(temp_array, len(temp_array)-3)
			item_note_info_list = split(temp_array, "~!~")
		End If
		set subNodesList = nothing

		set subNodesList = node.SelectNodes("/form/question[number='"& number &"']/itemAnsList")
		temp_array = ""
		If Not subNodesList Is Nothing Then
			for each nodeItem in subNodesList
				temp_array = temp_array & nodeItem.Text & "~!~"
			next
		End If
		If temp_array <> "" Then
			temp_array = left(temp_array, len(temp_array)-3)
			item_ans_list = split(temp_array, "~!~")
			If make_array_checkboxes = True Then
				for i = 0 to UBound(item_ans_list)
					If item_ans_list(i) = "1" Then item_ans_list(i) = checked
					If item_ans_list(i) = "0" Then item_ans_list(i) = unchecked
				next
			End If
		End If
		set subNodesList = nothing

		set subNodesList = node.SelectNodes("/form/question[number='"& number &"']/itemDetailList")
		temp_array = ""
		If Not subNodesList Is Nothing Then
			for each nodeItem in subNodesList
				temp_array = temp_array & nodeItem.Text & "~!~"
			next
		End If
		If temp_array <> "" Then
			temp_array = left(temp_array, len(temp_array)-3)
			item_detail_list = split(temp_array, "~!~")
		End If
		set subNodesList = nothing

		Set subNode = node.SelectSingleNode("allowPrefil")
		If Not subNode Is Nothing Then
			allow_prefil = subNode.Text
			If allow_prefil = "0" Then allow_prefil = False
			If allow_prefil = "-1" Then allow_prefil = True
		End If


		set subNodesList = node.SelectNodes("/form/question[number='"& number &"']/supplementalQuestions")
		temp_array = ""
		If Not subNodesList Is Nothing Then
			for each nodeItem in subNodesList
				' MsgBox number & vbCr & "Supplemental quesitons subnode text" & vbCr & vbCr & nodeItem.text
				temp_array = temp_array & nodeItem.Text & "~!~"
			next
		End If
		If temp_array <> "" Then
			temp_array = left(temp_array, len(temp_array)-3)
			supplemental_questions = split(temp_array, "~!~")
		End If
		set subNodesList = nothing

		Set subNode = node.SelectSingleNode("entirelyBlank")
		If Not subNode Is Nothing Then
			entirely_blank = subNode.Text
			If entirely_blank = "0" Then entirely_blank = False
			If entirely_blank = "-1" Then entirely_blank = True
		End If
		Set subNode = node.SelectSingleNode("detailArrayExists")
		If Not subNode Is Nothing Then
			detail_array_exists = subNode.Text
			If detail_array_exists = "0" Then detail_array_exists = False
			If detail_array_exists = "-1" Then detail_array_exists = True
		End If
		Set subNode = node.SelectSingleNode("detailSource")
		If Not subNode Is Nothing Then detail_source = subNode.Text


		set subNodesList = node.SelectNodes("/form/question[number='"& number &"']/detailInterviewNotes")
		temp_array = ""
		If Not subNodesList Is Nothing Then
			for each nodeItem in subNodesList
				temp_array = temp_array & nodeItem.Text & "~!~"
			next
		End If
		If temp_array <> "" Then
			temp_array = left(temp_array, len(temp_array)-3)
			If InStr(temp_array, "~!~") <> 0 Then detail_interview_notes = split(temp_array, "~!~")
			If InStr(temp_array, "~!~") = 0 Then detail_interview_notes = array(temp_array)
		End If
		set subNodesList = nothing

		set subNodesList = node.SelectNodes("/form/question[number='"& number &"']/detailWriteInInfo")
		temp_array = ""
		If Not subNodesList Is Nothing Then
			for each nodeItem in subNodesList
				temp_array = temp_array & nodeItem.Text & "~!~"
			next
		End If
		If temp_array <> "" Then
			temp_array = left(temp_array, len(temp_array)-3)
			If InStr(temp_array, "~!~") <> 0 Then detail_write_in_info = split(temp_array, "~!~")
			If InStr(temp_array, "~!~") = 0 Then detail_write_in_info = array(temp_array)
		End If
		set subNodesList = nothing

		set subNodesList = node.SelectNodes("/form/question[number='"& number &"']/detailVerifStatus")
		temp_array = ""
		If Not subNodesList Is Nothing Then
			for each nodeItem in subNodesList
				temp_array = temp_array & nodeItem.Text & "~!~"
			next
		End If
		If temp_array <> "" Then
			temp_array = left(temp_array, len(temp_array)-3)
			If InStr(temp_array, "~!~") <> 0 Then detail_verif_status = split(temp_array, "~!~")
			If InStr(temp_array, "~!~") = 0 Then detail_verif_status = array(temp_array)
		End If
		set subNodesList = nothing

		set subNodesList = node.SelectNodes("/form/question[number='"& number &"']/detailVerifNotes")
		temp_array = ""
		If Not subNodesList Is Nothing Then
			for each nodeItem in subNodesList
				temp_array = temp_array & nodeItem.Text & "~!~"
			next
		End If
		If temp_array <> "" Then
			temp_array = left(temp_array, len(temp_array)-3)
			If InStr(temp_array, "~!~") <> 0 Then detail_verif_notes = split(temp_array, "~!~")
			If InStr(temp_array, "~!~") = 0 Then detail_verif_notes = array(temp_array)
		End If
		set subNodesList = nothing

		set subNodesList = node.SelectNodes("/form/question[number='"& number &"']/detailEditBtn")
		temp_array = ""
		If Not subNodesList Is Nothing Then
			for each nodeItem in subNodesList
				temp_array = temp_array & nodeItem.Text & "~!~"
			next
		End If
		If temp_array <> "" Then
			temp_array = left(temp_array, len(temp_array)-3)
			If InStr(temp_array, "~!~") <> 0 Then detail_edit_btn = split(temp_array, "~!~")
			If InStr(temp_array, "~!~") = 0 Then detail_edit_btn = array(temp_array)
		End If
		set subNodesList = nothing

		set subNodesList = node.SelectNodes("/form/question[number='"& number &"']/detailResidentName")
		temp_array = ""
		If Not subNodesList Is Nothing Then
			for each nodeItem in subNodesList
				temp_array = temp_array & nodeItem.Text & "~!~"
			next
		End If
		If temp_array <> "" Then
			temp_array = left(temp_array, len(temp_array)-3)
			If InStr(temp_array, "~!~") <> 0 Then detail_resident_name = split(temp_array, "~!~")
			If InStr(temp_array, "~!~") = 0 Then detail_resident_name = array(temp_array)
		End If
		set subNodesList = nothing

		set subNodesList = node.SelectNodes("/form/question[number='"& number &"']/detailValue")
		temp_array = ""
		If Not subNodesList Is Nothing Then
			for each nodeItem in subNodesList
				temp_array = temp_array & nodeItem.Text & "~!~"
			next
		End If
		If temp_array <> "" Then
			temp_array = left(temp_array, len(temp_array)-3)
			If InStr(temp_array, "~!~") <> 0 Then detail_value = split(temp_array, "~!~")
			If InStr(temp_array, "~!~") = 0 Then detail_value = array(temp_array)
		End If
		set subNodesList = nothing

		set subNodesList = node.SelectNodes("/form/question[number='"& number &"']/detailType")
		temp_array = ""
		If Not subNodesList Is Nothing Then
			for each nodeItem in subNodesList
				temp_array = temp_array & nodeItem.Text & "~!~"
			next
		End If
		If temp_array <> "" Then
			temp_array = left(temp_array, len(temp_array)-3)
			If InStr(temp_array, "~!~") <> 0 Then detail_type = split(temp_array, "~!~")
			If InStr(temp_array, "~!~") = 0 Then detail_type = array(temp_array)
		End If
		set subNodesList = nothing

		set subNodesList = node.SelectNodes("/form/question[number='"& number &"']/detailHourlyWage")
		temp_array = ""
		If Not subNodesList Is Nothing Then
			for each nodeItem in subNodesList
				temp_array = temp_array & nodeItem.Text & "~!~"
			next
		End If
		If temp_array <> "" Then
			temp_array = left(temp_array, len(temp_array)-3)
			If InStr(temp_array, "~!~") <> 0 Then detail_hourly_wage = split(temp_array, "~!~")
			If InStr(temp_array, "~!~") = 0 Then detail_hourly_wage = array(temp_array)
		End If
		set subNodesList = nothing

		set subNodesList = node.SelectNodes("/form/question[number='"& number &"']/detailHoursPerWeek")
		temp_array = ""
		If Not subNodesList Is Nothing Then
			for each nodeItem in subNodesList
				temp_array = temp_array & nodeItem.Text & "~!~"
			next
		End If
		If temp_array <> "" Then
			temp_array = left(temp_array, len(temp_array)-3)
			If InStr(temp_array, "~!~") <> 0 Then detail_hours_per_week = split(temp_array, "~!~")
			If InStr(temp_array, "~!~") = 0 Then detail_hours_per_week = array(temp_array)
		End If
		set subNodesList = nothing

		set subNodesList = node.SelectNodes("/form/question[number='"& number &"']/detailBusiness")
		temp_array = ""
		If Not subNodesList Is Nothing Then
			for each nodeItem in subNodesList
				temp_array = temp_array & nodeItem.Text & "~!~"
			next
		End If
		If temp_array <> "" Then
			temp_array = left(temp_array, len(temp_array)-3)
			If InStr(temp_array, "~!~") <> 0 Then detail_business = split(temp_array, "~!~")
			If InStr(temp_array, "~!~") = 0 Then detail_business = array(temp_array)
		End If
		set subNodesList = nothing

		set subNodesList = node.SelectNodes("/form/question[number='"& number &"']/detailMonthlyAmount")
		temp_array = ""
		If Not subNodesList Is Nothing Then
			for each nodeItem in subNodesList
				temp_array = temp_array & nodeItem.Text & "~!~"
			next
		End If
		If temp_array <> "" Then
			temp_array = left(temp_array, len(temp_array)-3)
			If InStr(temp_array, "~!~") <> 0 Then detail_monthly_amount = split(temp_array, "~!~")
			If InStr(temp_array, "~!~") = 0 Then detail_monthly_amount = array(temp_array)
		End If
		set subNodesList = nothing

		set subNodesList = node.SelectNodes("/form/question[number='"& number &"']/detailDate")
		temp_array = ""
		If Not subNodesList Is Nothing Then
			for each nodeItem in subNodesList
				temp_array = temp_array & nodeItem.Text & "~!~"
			next
		End If
		If temp_array <> "" Then
			temp_array = left(temp_array, len(temp_array)-3)
			If InStr(temp_array, "~!~") <> 0 Then detail_date = split(temp_array, "~!~")
			If InStr(temp_array, "~!~") = 0 Then detail_date = array(temp_array)
		End If
		set subNodesList = nothing

		set subNodesList = node.SelectNodes("/form/question[number='"& number &"']/detailFrequency")
		temp_array = ""
		If Not subNodesList Is Nothing Then
			for each nodeItem in subNodesList
				temp_array = temp_array & nodeItem.Text & "~!~"
			next
		End If
		If temp_array <> "" Then
			temp_array = left(temp_array, len(temp_array)-3)
			If InStr(temp_array, "~!~") <> 0 Then detail_frequency = split(temp_array, "~!~")
			If InStr(temp_array, "~!~") = 0 Then detail_frequency = array(temp_array)
		End If
		set subNodesList = nothing

		set subNodesList = node.SelectNodes("/form/question[number='"& number &"']/detailAmount")
		temp_array = ""
		If Not subNodesList Is Nothing Then
			for each nodeItem in subNodesList
				temp_array = temp_array & nodeItem.Text & "~!~"
			next
		End If
		If temp_array <> "" Then
			temp_array = left(temp_array, len(temp_array)-3)
			If InStr(temp_array, "~!~") <> 0 Then detail_amount = split(temp_array, "~!~")
			If InStr(temp_array, "~!~") = 0 Then detail_amount = array(temp_array)
		End If
		set subNodesList = nothing

		set subNodesList = node.SelectNodes("/form/question[number='"& number &"']/detailCurrent")
		temp_array = ""
		If Not subNodesList Is Nothing Then
			for each nodeItem in subNodesList
				temp_array = temp_array & nodeItem.Text & "~!~"
			next
		End If
		If temp_array <> "" Then
			temp_array = left(temp_array, len(temp_array)-3)
			If InStr(temp_array, "~!~") <> 0 Then detail_current = split(temp_array, "~!~")
			If InStr(temp_array, "~!~") = 0 Then detail_current = array(temp_array)
		End If
		set subNodesList = nothing

		set subNodesList = node.SelectNodes("/form/question[number='"& number &"']/detailExplain")
		temp_array = ""
		If Not subNodesList Is Nothing Then
			for each nodeItem in subNodesList
				temp_array = temp_array & nodeItem.Text & "~!~"
			next
		End If
		If temp_array <> "" Then
			temp_array = left(temp_array, len(temp_array)-3)
			If InStr(temp_array, "~!~") <> 0 Then detail_explain = split(temp_array, "~!~")
			If InStr(temp_array, "~!~") = 0 Then detail_explain = array(temp_array)
		End If
		set subNodesList = nothing

		Set subNode = node.SelectSingleNode("detailButtonLabel")
		If Not subNode Is Nothing Then detail_button_label = subNode.Text
		Set subNode = node.SelectSingleNode("housingPayment")
		If Not subNode Is Nothing Then housing_payment = subNode.Text
		Set subNode = node.SelectSingleNode("heatAirCheckbox")
		If Not subNode Is Nothing Then
			heat_air_checkbox = subNode.Text
			If heat_air_checkbox <> "" Then heat_air_checkbox = heat_air_checkbox * 1
		End If
		Set subNode = node.SelectSingleNode("electricCheckbox")
		If Not subNode Is Nothing Then
			electric_checkbox = subNode.Text
			If electric_checkbox <> "" Then electric_checkbox = electric_checkbox * 1
		End If
		Set subNode = node.SelectSingleNode("phoneCheckbox")
		If Not subNode Is Nothing Then
			phone_checkbox = subNode.Text
			If phone_checkbox <> "" Then phone_checkbox = phone_checkbox * 1
		End If
		Set subNode = node.SelectSingleNode("subsidyYN")
		If Not subNode Is Nothing Then subsidy_yn = subNode.Text
		Set subNode = node.SelectSingleNode("subsidyAmount")
		If Not subNode Is Nothing Then subsidy_amount = subNode.Text
		Set subNode = node.SelectSingleNode("verifStatus")
		If Not subNode Is Nothing Then verif_status = subNode.Text
		Set subNode = node.SelectSingleNode("verifNotes")
		If Not subNode Is Nothing Then verif_notes = subNode.Text
		Set subNode = node.SelectSingleNode("guideBtn")
		If Not subNode Is Nothing Then
			guide_btn = subNode.Text
			guide_btn = guide_btn * 1
		End If
		Set subNode = node.SelectSingleNode("verifBtn")
		If Not subNode Is Nothing Then
			verif_btn = verif_btn = verif_btn * 1
		End If
		Set subNode = node.SelectSingleNode("prefilBtn")
		If Not subNode Is Nothing Then
			prefil_btn = subNode.Text
			If prefil_btn <> "" Then prefil_btn = prefil_btn * 1
		End If
		Set subNode = node.SelectSingleNode("addToArrayBtn")
		If Not subNode Is Nothing Then
			add_to_array_btn = subNode.Text
			If add_to_array_btn <> "" Then add_to_array_btn = add_to_array_btn * 1
		End If
		Set subNode = node.SelectSingleNode("editInArrayBtn")
		If Not subNode Is Nothing Then
			edit_in_array_btn = subNode.Text
			If edit_in_array_btn <> "" Then edit_in_array_btn = edit_in_array_btn * 1
		End If

		Set subNode = node.SelectSingleNode("errMandated")
		If Not subNode Is Nothing Then
			mandated = subNode.Text
			If mandated = "0" Then mandated = False
			If mandated = "-1" Then mandated = True
		End If

		Set subNode = node.SelectSingleNode("errorInfo")
		If Not subNode Is Nothing Then error_info = subNode.Text

		Set subNode = node.SelectSingleNode("errorVerbiage")
		If Not subNode Is Nothing Then error_verbiage = subNode.Text

		Set subNode = node.SelectSingleNode("dialogPageNumb")
		If Not subNode Is Nothing Then
			dialog_page_numb = subNode.Text
			dialog_page_numb = dialog_page_numb * 1
		End If
		Set subNode = node.SelectSingleNode("dialogOrder")
		If Not subNode Is Nothing Then dialog_order = subNode.Text
		Set subNode = node.SelectSingleNode("dialogHeight")
		If Not subNode Is Nothing Then
			dialog_height = subNode.Text
			dialog_height = dialog_height * 1
		End If
	end sub

	public sub save_answer(root)
		Set quest = xmlDoc.createElement("question")
		root.appendChild quest

		Set XML_number = xmlDoc.createElement("number")
		quest.appendChild XML_number
		Set info_number = xmlDoc.createTextNode(number)
		XML_number.appendChild info_number

		Set XML_dialog_phrasing = xmlDoc.createElement("dialogPhrasing")
		quest.appendChild XML_dialog_phrasing
		Set info_number = xmlDoc.createTextNode(dialog_phrasing)
		XML_dialog_phrasing.appendChild info_number

		Set XML_note_phrasing = xmlDoc.createElement("notePhrasing")
		quest.appendChild XML_note_phrasing
		Set info_number = xmlDoc.createTextNode(note_phrasing)
		XML_note_phrasing.appendChild info_number

		Set XML_doc_phrasing = xmlDoc.createElement("docPhrasing")
		quest.appendChild XML_doc_phrasing
		Set info_number = xmlDoc.createTextNode(doc_phrasing)
		XML_doc_phrasing.appendChild info_number

		Set XML_sub_number = xmlDoc.createElement("subNumber")
		quest.appendChild XML_sub_number
		Set info_number = xmlDoc.createTextNode(sub_number)
		XML_sub_number.appendChild info_number

		Set XML_sub_phrase = xmlDoc.createElement("subPhrase")
		quest.appendChild XML_sub_phrase
		Set info_number = xmlDoc.createTextNode(sub_phrase)
		XML_sub_phrase.appendChild info_number

		Set XML_sub_note_phrase = xmlDoc.createElement("subNotePhrase")
		quest.appendChild XML_sub_note_phrase
		Set info_number = xmlDoc.createTextNode(sub_note_phrase)
		XML_sub_note_phrase.appendChild info_number

		Set XML_sub_answer = xmlDoc.createElement("subAnswer")
		quest.appendChild XML_sub_answer
		Set info_number = xmlDoc.createTextNode(sub_answer)
		XML_sub_answer.appendChild info_number


		Set XML_info_type = xmlDoc.createElement("infoType")
		quest.appendChild XML_info_type
		Set info_number = xmlDoc.createTextNode(info_type)
		XML_info_type.appendChild info_number

		Set XML_caf_answer = xmlDoc.createElement("cafAnswer")
		quest.appendChild XML_caf_answer
		Set info_number = xmlDoc.createTextNode(caf_answer)
		XML_caf_answer.appendChild info_number

		Set XML_answer_is_array = xmlDoc.createElement("answerIsArray")
		quest.appendChild XML_answer_is_array
		Set info_number = xmlDoc.createTextNode(answer_is_array)
		XML_answer_is_array.appendChild info_number

		Set XML_make_array_checkboxes = xmlDoc.createElement("makeArrayCheckboxes")
		quest.appendChild XML_make_array_checkboxes
		Set info_number = xmlDoc.createTextNode(make_array_checkboxes)
		XML_make_array_checkboxes.appendChild info_number

		Set XML_write_in_info = xmlDoc.createElement("writeInInfo")
		quest.appendChild XML_write_in_info
		Set info_number = xmlDoc.createTextNode(write_in_info)
		XML_write_in_info.appendChild info_number

		Set XML_interview_notes = xmlDoc.createElement("interviewNotes")
		quest.appendChild XML_interview_notes
		Set info_number = xmlDoc.createTextNode(interview_notes)
		XML_interview_notes.appendChild info_number

		If IsArray(item_info_list) = True Then
			for each cow in item_info_list
				Set XML_item_info_list = xmlDoc.createElement("itemInfoList")
				quest.appendChild XML_item_info_list
				Set info_number = xmlDoc.createTextNode(cow)
				XML_item_info_list.appendChild info_number
			Next
		End If

		If IsArray(item_note_info_list) = True Then
			for each cow in item_note_info_list
				Set XML_item_note_info_list = xmlDoc.createElement("itemNoteInfoList")
				quest.appendChild XML_item_note_info_list
				Set info_number = xmlDoc.createTextNode(cow)
				XML_item_note_info_list.appendChild info_number
			Next
		End If

		If IsArray(item_ans_list) = True Then
			for each cow in item_ans_list
				Set XML_item_ans_list = xmlDoc.createElement("itemAnsList")
				quest.appendChild XML_item_ans_list
				Set info_number = xmlDoc.createTextNode(cow)
				XML_item_ans_list.appendChild info_number
			Next
		End If

		If IsArray(item_detail_list) = True Then
			for each cow in item_detail_list
				Set XML_item_detail_list = xmlDoc.createElement("itemDetailList")
				quest.appendChild XML_item_detail_list
				Set info_number = xmlDoc.createTextNode(cow)
				XML_item_detail_list.appendChild info_number
			Next
		End If

		Set XML_allow_prefil = xmlDoc.createElement("allowPrefil")
		quest.appendChild XML_allow_prefil
		Set info_number = xmlDoc.createTextNode(allow_prefil)
		XML_allow_prefil.appendChild info_number

		If IsArray(supplemental_questions) = True Then
			for each cow in supplemental_questions
				Set XML_supplemental_questions = xmlDoc.createElement("supplementalQuestions")
				quest.appendChild XML_supplemental_questions
				Set info_number = xmlDoc.createTextNode(cow)
				XML_supplemental_questions.appendChild info_number
			Next
		End If

		Set XML_entirely_blank = xmlDoc.createElement("entirelyBlank")
		quest.appendChild XML_entirely_blank
		Set info_number = xmlDoc.createTextNode(entirely_blank)
		XML_entirely_blank.appendChild info_number

		Set XML_detail_array_exists = xmlDoc.createElement("detailArrayExists")
		quest.appendChild XML_detail_array_exists
		Set info_number = xmlDoc.createTextNode(detail_array_exists)
		XML_detail_array_exists.appendChild info_number

		Set XML_detail_source = xmlDoc.createElement("detailSource")
		quest.appendChild XML_detail_source
		Set info_number = xmlDoc.createTextNode(detail_source)
		XML_detail_source.appendChild info_number

		If IsArray(detail_interview_notes) = True Then
			for each cow in detail_interview_notes
				Set XML_detail_interview_notes = xmlDoc.createElement("detailInterviewNotes")
				quest.appendChild XML_detail_interview_notes
				Set info_number = xmlDoc.createTextNode(cow)
				XML_detail_interview_notes.appendChild info_number
			Next
		End If

		If IsArray(detail_write_in_info) = True Then
			for each cow in detail_write_in_info
				Set XML_detail_write_in_info = xmlDoc.createElement("detailWriteInInfo")
				quest.appendChild XML_detail_write_in_info
				Set info_number = xmlDoc.createTextNode(cow)
				XML_detail_write_in_info.appendChild info_number
			Next
		End If

		If IsArray(detail_verif_status) = True Then
			for each cow in detail_verif_status
				Set XML_detail_verif_status = xmlDoc.createElement("detailVerifStatus")
				quest.appendChild XML_detail_verif_status
				Set info_number = xmlDoc.createTextNode(cow)
				XML_detail_verif_status.appendChild info_number
			Next
		End If

		If IsArray(detail_verif_notes) = True Then
			for each cow in detail_verif_notes
				Set XML_detail_verif_notes = xmlDoc.createElement("detailVerifNotes")
				quest.appendChild XML_detail_verif_notes
				Set info_number = xmlDoc.createTextNode(cow)
				XML_detail_verif_notes.appendChild info_number
			Next
		End If

		If IsArray(detail_edit_btn) = True Then
			for each cow in detail_edit_btn
				Set XML_detail_edit_btn = xmlDoc.createElement("detailEditBtn")
				quest.appendChild XML_detail_edit_btn
				Set info_number = xmlDoc.createTextNode(cow)
				XML_detail_edit_btn.appendChild info_number
			Next
		End If

		If IsArray(detail_resident_name) = True Then
			for each cow in detail_resident_name
				Set XML_detail_resident_name = xmlDoc.createElement("detailResidentName")
				quest.appendChild XML_detail_resident_name
				Set info_number = xmlDoc.createTextNode(cow)
				XML_detail_resident_name.appendChild info_number
			Next
		End If

		If IsArray(detail_value) = True Then
			for each cow in detail_value
				Set XML_detail_value = xmlDoc.createElement("detailValue")
				quest.appendChild XML_detail_value
				Set info_number = xmlDoc.createTextNode(cow)
				XML_detail_value.appendChild info_number
			Next
		End If

		If IsArray(detail_type) = True Then
			for each cow in detail_type
				Set XML_detail_type = xmlDoc.createElement("detailType")
				quest.appendChild XML_detail_type
				Set info_number = xmlDoc.createTextNode(cow)
				XML_detail_type.appendChild info_number
			Next
		End If

		If IsArray(detail_hourly_wage) = True Then
			for each cow in detail_hourly_wage
				Set XML_detail_hourly_wage = xmlDoc.createElement("detailHourlyWage")
				quest.appendChild XML_detail_hourly_wage
				Set info_number = xmlDoc.createTextNode(cow)
				XML_detail_hourly_wage.appendChild info_number
			Next
		End If

		If IsArray(detail_hours_per_week) = True Then
			for each cow in detail_hours_per_week
				Set XML_detail_hours_per_week = xmlDoc.createElement("detailHoursPerWeek")
				quest.appendChild XML_detail_hours_per_week
				Set info_number = xmlDoc.createTextNode(cow)
				XML_detail_hours_per_week.appendChild info_number
			Next
		End If

		If IsArray(detail_business) = True Then
			for each cow in detail_business
				Set XML_detail_business = xmlDoc.createElement("detailBusiness")
				quest.appendChild XML_detail_business
				Set info_number = xmlDoc.createTextNode(cow)
				XML_detail_business.appendChild info_number
			Next
		End If

		If IsArray(detail_monthly_amount) = True Then
			for each cow in detail_monthly_amount
				Set XML_detail_monthly_amount = xmlDoc.createElement("detailMonthlyAmount")
				quest.appendChild XML_detail_monthly_amount
				Set info_number = xmlDoc.createTextNode(cow)
				XML_detail_monthly_amount.appendChild info_number
			Next
		End If

		If IsArray(detail_date) = True Then
			for each cow in detail_date
				Set XML_detail_date = xmlDoc.createElement("detailDate")
				quest.appendChild XML_detail_date
				Set info_number = xmlDoc.createTextNode(cow)
				XML_detail_date.appendChild info_number
			Next
		End If

		If IsArray(detail_frequency) = True Then
			for each cow in detail_frequency
				Set XML_detail_frequency = xmlDoc.createElement("detailFrequency")
				quest.appendChild XML_detail_frequency
				Set info_number = xmlDoc.createTextNode(cow)
				XML_detail_frequency.appendChild info_number
			Next
		End If

		If IsArray(detail_amount) = True Then
			for each cow in detail_amount
				Set XML_detail_amount = xmlDoc.createElement("detailAmount")
				quest.appendChild XML_detail_amount
				Set info_number = xmlDoc.createTextNode(cow)
				XML_detail_amount.appendChild info_number
			Next
		End If

		If IsArray(detail_current) = True Then
			for each cow in detail_current
				Set XML_detail_current = xmlDoc.createElement("detailCurrent")
				quest.appendChild XML_detail_current
				Set info_number = xmlDoc.createTextNode(cow)
				XML_detail_current.appendChild info_number
			Next
		End If

		If IsArray(detail_explain) = True Then
			for each cow in detail_explain
				Set XML_detail_explain = xmlDoc.createElement("detailExplain")
				quest.appendChild XML_detail_explain
				Set info_number = xmlDoc.createTextNode(cow)
				XML_detail_explain.appendChild info_number
			Next
		End If

		Set XML_detail_button_label = xmlDoc.createElement("detailButtonLabel")
		quest.appendChild XML_detail_button_label
		Set info_number = xmlDoc.createTextNode(detail_button_label)
		XML_detail_button_label.appendChild info_number

		Set XML_housing_payment = xmlDoc.createElement("housingPayment")
		quest.appendChild XML_housing_payment
		Set info_number = xmlDoc.createTextNode(housing_payment)
		XML_housing_payment.appendChild info_number

		Set XML_heat_air_checkbox = xmlDoc.createElement("heatAirCheckbox")
		quest.appendChild XML_heat_air_checkbox
		Set info_number = xmlDoc.createTextNode(heat_air_checkbox)
		XML_heat_air_checkbox.appendChild info_number

		Set XML_electric_checkbox= xmlDoc.createElement("electricCheckbox")
		quest.appendChild XML_electric_checkbox
		Set info_number = xmlDoc.createTextNode(electric_checkbox)
		XML_electric_checkbox.appendChild info_number

		Set XML_phone_checkbox = xmlDoc.createElement("phoneCheckbox")
		quest.appendChild XML_phone_checkbox
		Set info_number = xmlDoc.createTextNode(phone_checkbox)
		XML_phone_checkbox.appendChild info_number

		Set XML_subsidy_yn = xmlDoc.createElement("subsidyYN")
		quest.appendChild XML_subsidy_yn
		Set info_number = xmlDoc.createTextNode(subsidy_yn)
		XML_subsidy_yn.appendChild info_number

		Set XML_subsidy_amount= xmlDoc.createElement("subsidyAmount")
		quest.appendChild XML_subsidy_amount
		Set info_number = xmlDoc.createTextNode(subsidy_amount)
		XML_subsidy_amount.appendChild info_number

		Set XML_verif_status = xmlDoc.createElement("verifStatus")
		quest.appendChild XML_verif_status
		Set info_number = xmlDoc.createTextNode(verif_status)
		XML_verif_status.appendChild info_number

		Set XML_verif_notes = xmlDoc.createElement("verifNotes")
		quest.appendChild XML_verif_notes
		Set info_number = xmlDoc.createTextNode(verif_notes)
		XML_verif_notes.appendChild info_number

		Set XML_guide_btn = xmlDoc.createElement("guideBtn")
		quest.appendChild XML_guide_btn
		Set info_number = xmlDoc.createTextNode(guide_btn)
		XML_guide_btn.appendChild info_number

		Set XML_verif_btn = xmlDoc.createElement("verifBtn")
		quest.appendChild XML_verif_btn
		Set info_number = xmlDoc.createTextNode(verif_btn)
		XML_verif_btn.appendChild info_number

		Set XML_prefil_btn = xmlDoc.createElement("prefilBtn")
		quest.appendChild XML_prefil_btn
		Set info_number = xmlDoc.createTextNode(prefil_btn)
		XML_prefil_btn.appendChild info_number

		Set XML_add_to_array_btn = xmlDoc.createElement("addToArrayBtn")
		quest.appendChild XML_add_to_array_btn
		Set info_number = xmlDoc.createTextNode(add_to_array_btn)
		XML_add_to_array_btn.appendChild info_number

		Set XML_edit_in_array_btn = xmlDoc.createElement("editInArrayBtn")
		quest.appendChild XML_edit_in_array_btn
		Set info_number = xmlDoc.createTextNode(edit_in_array_btn)
		XML_edit_in_array_btn.appendChild info_number

		Set XML_error_mandated = xmlDoc.createElement("errMandated")
		quest.appendChild XML_error_mandated
		Set info_number = xmlDoc.createTextNode(mandated)
		XML_error_mandated.appendChild info_number

		Set XML_error_info = xmlDoc.createElement("errorInfo")
		quest.appendChild XML_error_info
		Set info_number = xmlDoc.createTextNode(error_info)
		XML_error_info.appendChild info_number

		Set XML_error_verbiage = xmlDoc.createElement("errorVerbiage")
		quest.appendChild XML_error_verbiage
		Set info_number = xmlDoc.createTextNode(error_verbiage)
		XML_error_verbiage.appendChild info_number

		Set XML_dialog_page_numb = xmlDoc.createElement("dialogPageNumb")
		quest.appendChild XML_dialog_page_numb
		Set info_number = xmlDoc.createTextNode(dialog_page_numb)
		XML_dialog_page_numb.appendChild info_number

		Set XML_dialog_order = xmlDoc.createElement("dialogOrder")
		quest.appendChild XML_dialog_order
		Set info_number = xmlDoc.createTextNode(dialog_order)
		XML_dialog_order.appendChild info_number

		Set XML_dialog_height = xmlDoc.createElement("dialogHeight")
		quest.appendChild XML_dialog_height
		Set info_number = xmlDoc.createTextNode(dialog_height)
		XML_dialog_height.appendChild info_number
	end sub
end class


function array_details_dlg(form_question_numb, selected)
	det_dlg_len = 175
	dialog_title = ""
	temp_interview_notes = FORM_QUESTION_ARRAY(form_question_numb).detail_interview_notes(selected)
	temp_write_in_info = FORM_QUESTION_ARRAY(form_question_numb).detail_write_in_info(selected)
	' MsgBox "Question - " & FORM_QUESTION_ARRAY(form_question_numb).number & vbCr & "Source - " & FORM_QUESTION_ARRAY(form_question_numb).detail_source
	If FORM_QUESTION_ARRAY(form_question_numb).detail_source = "jobs" Then
		temp_resident_name = FORM_QUESTION_ARRAY(form_question_numb).detail_resident_name(selected)
		temp_hourly_wage = FORM_QUESTION_ARRAY(form_question_numb).detail_hourly_wage(selected)
		temp_business = FORM_QUESTION_ARRAY(form_question_numb).detail_business(selected)
		temp_monthly_amount = FORM_QUESTION_ARRAY(form_question_numb).detail_monthly_amount(selected)
		temp_hours_per_week = FORM_QUESTION_ARRAY(form_question_numb).detail_hours_per_week(selected)
		dialog_title = "Job Details"
		instruction_text = "Enter Job Details/Information"
	ElseIf FORM_QUESTION_ARRAY(form_question_numb).detail_source = "assets" Then
		temp_resident_name = FORM_QUESTION_ARRAY(form_question_numb).detail_resident_name(selected)
		temp_type = FORM_QUESTION_ARRAY(form_question_numb).detail_type(selected)
		temp_value = FORM_QUESTION_ARRAY(form_question_numb).detail_value(selected)
		temp_explain = FORM_QUESTION_ARRAY(form_question_numb).detail_explain(selected)
		dialog_title = "Asset Details"
		instruction_text = "Enter Details of Household Assets"
	ElseIf FORM_QUESTION_ARRAY(form_question_numb).detail_source = "unea" Then
		temp_resident_name = FORM_QUESTION_ARRAY(form_question_numb).detail_resident_name(selected)
		temp_type = FORM_QUESTION_ARRAY(form_question_numb).detail_type(selected)
		temp_date = FORM_QUESTION_ARRAY(form_question_numb).detail_date(selected)
		temp_amount = FORM_QUESTION_ARRAY(form_question_numb).detail_amount(selected)
		temp_frequency = FORM_QUESTION_ARRAY(form_question_numb).detail_frequency(selected)
		dialog_title = "UNEA Details"
		instruction_text = "Enter Details of Income from Unearned Sources"
	ElseIf FORM_QUESTION_ARRAY(form_question_numb).detail_source = "shel-hest" Then
		det_dlg_len = 235
		temp_type = FORM_QUESTION_ARRAY(form_question_numb).detail_type(selected)
		temp_amount = FORM_QUESTION_ARRAY(form_question_numb).detail_amount(selected)
		temp_frequency = FORM_QUESTION_ARRAY(form_question_numb).detail_frequency(selected)
		dialog_title = "Housing and Utilities Details"
		instruction_text = "Enter Detail of Changes to Housing Expenses"

		temp_housing_payment = FORM_QUESTION_ARRAY(form_question_numb).housing_payment
		temp_heat_ac = FORM_QUESTION_ARRAY(form_question_numb).heat_air_checkbox
		temp_electric = FORM_QUESTION_ARRAY(form_question_numb).electric_checkbox
		temp_phone = FORM_QUESTION_ARRAY(form_question_numb).phone_checkbox
		temp_sub_yn = FORM_QUESTION_ARRAY(form_question_numb).subsidy_yn
		temp_sub_amt = FORM_QUESTION_ARRAY(form_question_numb).subsidy_amount
	ElseIf FORM_QUESTION_ARRAY(form_question_numb).detail_source = "expense" Then
		temp_resident_name = FORM_QUESTION_ARRAY(form_question_numb).detail_resident_name(selected)
		temp_amount = FORM_QUESTION_ARRAY(form_question_numb).detail_amount(selected)
		temp_current = FORM_QUESTION_ARRAY(form_question_numb).detail_current(selected)
		dialog_title = "Expense Details"
		instruction_text = "Enter Details of Expenses"
	ElseIf FORM_QUESTION_ARRAY(form_question_numb).detail_source = "winnings" Then
		temp_resident_name = FORM_QUESTION_ARRAY(form_question_numb).detail_resident_name(selected)
		temp_amount = FORM_QUESTION_ARRAY(form_question_numb).detail_amount(selected)
		temp_date = FORM_QUESTION_ARRAY(form_question_numb).detail_date(selected)
		dialog_title = "Expense Details"
		instruction_text = "Enter Details of Expenses"
	ElseIf FORM_QUESTION_ARRAY(form_question_numb).detail_source = "changes" Then
		temp_resident_name = FORM_QUESTION_ARRAY(form_question_numb).detail_resident_name(selected)
		temp_date = FORM_QUESTION_ARRAY(form_question_numb).detail_date(selected)
		temp_explain = FORM_QUESTION_ARRAY(form_question_numb).detail_explain(selected)
		dialog_title = "Changes Details"
		instruction_text = "Enter Details of Reported Changes"
	End If

	Do
		'TODO - Add a way to indicate this was ONLY a Verbal report
		Dialog1 = ""
		BeginDialog Dialog1, 0, 0, 321, det_dlg_len, dialog_title & ""
			Text 10, 5, 250, 20, FORM_QUESTION_ARRAY(form_question_numb).doc_phrasing
			Text 10, 20, 250, 10, instruction_text

			If FORM_QUESTION_ARRAY(form_question_numb).detail_source = "jobs" Then
				Text 10, 35, 70, 10, "EMPLOYEE NAME:"
				ComboBox 10, 45, 135, 45, pick_a_client+chr(9)+""+chr(9)+temp_resident_name, temp_resident_name
				Text 150, 35, 60, 10, "HOURLY WAGE:"
				EditBox 150, 45, 60, 15, temp_hourly_wage
				Text 215, 35, 105, 10, "GROSS MONTHLY EARNINGS:"
				EditBox 215, 45, 100, 15, temp_monthly_amount
				Text 10, 65, 60, 10, "HOURS/WEEK:"
				EditBox 10, 75, 60, 15, temp_hours_per_week
				Text 75, 65, 105, 10, "EMPLOYER/BUSINESS NAME:"
				EditBox 75, 75, 240, 15, temp_business
			ElseIf FORM_QUESTION_ARRAY(form_question_numb).detail_source = "assets" Then
				Text 10, 35, 70, 10, "OWNER(S) NAME:"
				ComboBox 10, 45, 150, 45, pick_a_client+chr(9)+""+chr(9)+temp_resident_name, temp_resident_name
				Text 165, 35, 75, 10, "TYPE:"
				EditBox 165, 45, 150, 15, temp_type
				Text 10, 65, 195, 10, "INFO (Account Number or Institution):"
				EditBox 10, 75, 195, 15, temp_explain
				Text 210, 65, 105, 10, "VALUE OR AMOUNT:"
				EditBox 210, 75, 105, 15, temp_value
			ElseIf FORM_QUESTION_ARRAY(form_question_numb).detail_source = "unea" Then
				Text 10, 35, 70, 10, "NAME:"
				ComboBox 10, 45, 135, 45, pick_a_client+chr(9)+""+chr(9)+temp_resident_name, temp_resident_name
				Text 150, 35, 60, 10, "TYPE:"
				EditBox 150, 45, 165, 15, temp_type
				Text 10, 65, 55, 10, "START DATE:"
				EditBox 10, 75, 100, 15, temp_date
				Text 115, 65, 55, 10, "AMOUNT:"
				EditBox 115, 75, 100, 15, temp_amount
				Text 220, 65, 95, 10, "FREQUENCY:"
				DropListBox 220, 75, 95, 15, ""+chr(9)+"Weekly"+chr(9)+"Bi-weekly"+chr(9)+"Semi-monthly"+chr(9)+"Once a month", temp_frequency
			ElseIf FORM_QUESTION_ARRAY(form_question_numb).detail_source = "shel-hest" Then
				Text 10, 35, 70, 10, "EXPENSE:"
				EditBox 10, 45, 305, 15, temp_type
				Text 10, 65, 75, 10, "AMOUNT:"
				EditBox 10, 75, 195, 15, temp_amount
				Text 210, 65, 105, 10, "HOW OFTEN?"
				EditBox 210, 75, 105, 15, temp_frequency

				Text 10, 95, 125, 10, "RENT/MORTGAGE PAYMENT:"
				EditBox 10, 105, 125, 15, temp_housing_payment
				Text 150, 95, 100, 10, "CHECK THE UTILITIES PAID:"
				Checkbox 150, 110, 50, 10, "Heat/AC", temp_heat_ac
				Checkbox 200, 110, 50, 10, "Electricity", temp_electric
				Checkbox 250, 110, 50, 10, "Telephone", temp_phone
				Text 10, 125, 80, 10, "SUBSIDY?"
				DropListBox 10, 135, 95, 15, ""+chr(9)+"No"+chr(9)+"Yes",temp_sub_yn
				Text 110, 125, 150, 10, "SUBSIDY AMOUNT"
				EditBox 110, 135, 205, 15, temp_sub_amt

			ElseIf FORM_QUESTION_ARRAY(form_question_numb).detail_source = "expense" Then
				Text 10, 35, 150, 10, "NAME OF PERSON PAYING:"
				ComboBox 10, 45, 305, 45, pick_a_client+chr(9)+""+chr(9)+temp_resident_name, temp_resident_name
				Text 10, 65, 100, 10, "MONTHLY AMOUNT:"
				EditBox 10, 75, 190, 15, temp_amount
				Text 205, 65, 110, 10, "CURRENTLY PAYING:"
				DropListBox 205, 75, 110, 15, ""+chr(9)+"No"+chr(9)+"Yes", temp_value
			ElseIf FORM_QUESTION_ARRAY(form_question_numb).detail_source = "winnings" Then
				Text 10, 35, 150, 10, "NAME OF PERSON WHO WON:"
				ComboBox 10, 45, 305, 45, pick_a_client+chr(9)+""+chr(9)+temp_resident_name, temp_resident_name
				Text 10, 65, 100, 10, "AMOUNT:"
				EditBox 10, 75, 190, 15, temp_amount
				Text 205, 65, 110, 10, "WIN DATE:"
				EditBox 205, 75, 110, 15, temp_date
			ElseIf FORM_QUESTION_ARRAY(form_question_numb).detail_source = "changes" Then
				Text 10, 35, 70, 10, "WHO?"
				ComboBox 10, 45, 200, 45, pick_a_client+chr(9)+""+chr(9)+temp_resident_name, temp_resident_name
				Text 215, 35, 120, 10, "DATE OF CHANGE:"
				EditBox 215, 45, 100, 15, temp_date
				Text 10, 65, 150, 10, "EXPLAIN THE CHANGE:"
				EditBox 10, 75, 305, 15, temp_explain
			End If


			Text 10, det_dlg_len-80, 110, 10, "CAF WRITE-IN INFORMATION:"
			EditBox 10, det_dlg_len-70, 305, 15, temp_write_in_info
			Text 10, det_dlg_len-50, 85, 10, "INTERVIEW NOTES:"
			EditBox 10, det_dlg_len-40, 305, 15, temp_interview_notes

			ButtonGroup ButtonPressed
				PushButton 265, det_dlg_len-20, 50, 15, "Return", return_btn
				PushButton 120, det_dlg_len-15, 75, 10, "ADD VERIFICATION", add_verif_jobs_btn
				PushButton 265, 10, 50, 10, "CLEAR", clear_job_btn
			Text 10, det_dlg_len-15, 110, 10, "Verification - " & FORM_QUESTION_ARRAY(form_question_numb).detail_verif_status(selected)
		EndDialog


		dialog Dialog1


		FORM_QUESTION_ARRAY(form_question_numb).detail_interview_notes(selected) = temp_interview_notes
		FORM_QUESTION_ARRAY(form_question_numb).detail_write_in_info(selected) = temp_write_in_info
		If FORM_QUESTION_ARRAY(form_question_numb).detail_source = "jobs" Then
			FORM_QUESTION_ARRAY(form_question_numb).detail_resident_name(selected) = temp_resident_name
			FORM_QUESTION_ARRAY(form_question_numb).detail_hourly_wage(selected) = temp_hourly_wage
			FORM_QUESTION_ARRAY(form_question_numb).detail_business(selected) = temp_business
			FORM_QUESTION_ARRAY(form_question_numb).detail_monthly_amount(selected) = temp_monthly_amount
			FORM_QUESTION_ARRAY(form_question_numb).detail_hours_per_week(selected) = temp_hours_per_week
		ElseIf FORM_QUESTION_ARRAY(form_question_numb).detail_source = "assets" Then
			FORM_QUESTION_ARRAY(form_question_numb).detail_resident_name(selected) = temp_resident_name
			FORM_QUESTION_ARRAY(form_question_numb).detail_type(selected) = temp_type
			FORM_QUESTION_ARRAY(form_question_numb).detail_value(selected) = temp_value
			FORM_QUESTION_ARRAY(form_question_numb).detail_explain(selected) = temp_explain
		ElseIf FORM_QUESTION_ARRAY(form_question_numb).detail_source = "unea" Then
			FORM_QUESTION_ARRAY(form_question_numb).detail_resident_name(selected) = temp_resident_name
			FORM_QUESTION_ARRAY(form_question_numb).detail_type(selected) = temp_type
			FORM_QUESTION_ARRAY(form_question_numb).detail_date(selected) = temp_date
			FORM_QUESTION_ARRAY(form_question_numb).detail_amount(selected) = temp_amount
			FORM_QUESTION_ARRAY(form_question_numb).detail_frequency(selected) = temp_frequency
		ElseIf FORM_QUESTION_ARRAY(form_question_numb).detail_source = "shel-hest" Then
			FORM_QUESTION_ARRAY(form_question_numb).detail_type(selected) = temp_type
			FORM_QUESTION_ARRAY(form_question_numb).detail_amount(selected) = temp_amount
			FORM_QUESTION_ARRAY(form_question_numb).detail_frequency(selected) = temp_frequency
			FORM_QUESTION_ARRAY(form_question_numb).housing_payment = temp_housing_payment
			FORM_QUESTION_ARRAY(form_question_numb).heat_air_checkbox = temp_heat_ac
			FORM_QUESTION_ARRAY(form_question_numb).electric_checkbox = temp_electric
			FORM_QUESTION_ARRAY(form_question_numb).phone_checkbox = temp_phone
			FORM_QUESTION_ARRAY(form_question_numb).subsidy_yn = temp_sub_yn
			FORM_QUESTION_ARRAY(form_question_numb).subsidy_amount	 = temp_sub_amt
		ElseIf FORM_QUESTION_ARRAY(form_question_numb).detail_source = "expense" Then
			FORM_QUESTION_ARRAY(form_question_numb).detail_resident_name(selected) = temp_resident_name
			FORM_QUESTION_ARRAY(form_question_numb).detail_amount(selected) = temp_amount
			FORM_QUESTION_ARRAY(form_question_numb).detail_current(selected) = temp_current
		ElseIf FORM_QUESTION_ARRAY(form_question_numb).detail_source = "winnings" Then
			FORM_QUESTION_ARRAY(form_question_numb).detail_resident_name(selected) = temp_resident_name
			FORM_QUESTION_ARRAY(form_question_numb).detail_amount(selected) = temp_amount
			FORM_QUESTION_ARRAY(form_question_numb).detail_date(selected) = temp_date
		ElseIf FORM_QUESTION_ARRAY(form_question_numb).detail_source = "changes" Then
			FORM_QUESTION_ARRAY(form_question_numb).detail_resident_name(selected) = temp_resident_name
			FORM_QUESTION_ARRAY(form_question_numb).detail_date(selected) = temp_date
			FORM_QUESTION_ARRAY(form_question_numb).detail_explain(selected) = temp_explain
		End If
		If FORM_QUESTION_ARRAY(form_question_numb).detail_source <> "shel-hest" Then
			If FORM_QUESTION_ARRAY(form_question_numb).detail_resident_name(selected) = "Select One..." Then FORM_QUESTION_ARRAY(form_question_numb).detail_resident_name(selected) = ""
		End If

		If ButtonPressed = -1 Then ButtonPressed = return_btn
		If ButtonPressed = add_verif_jobs_btn Then Call FORM_QUESTION_ARRAY(form_question_numb).capture_array_verif_detail(selected)
		If ButtonPressed = clear_job_btn Then
			If FORM_QUESTION_ARRAY(form_question_numb).detail_source = "jobs" Then
				temp_resident_name = ""
				temp_hourly_wage = ""
				temp_business = ""
				temp_monthly_amount = ""
				temp_hours_per_week = ""
			ElseIf FORM_QUESTION_ARRAY(form_question_numb).detail_source = "assets" Then
				temp_resident_name = ""
				temp_type = ""
				temp_value = ""
				temp_explain = ""
			ElseIf FORM_QUESTION_ARRAY(form_question_numb).detail_source = "unea" Then
				temp_resident_name = ""
				temp_type = ""
				temp_date = ""
				temp_amount = ""
				temp_frequency = ""
			ElseIf FORM_QUESTION_ARRAY(form_question_numb).detail_source = "shel-hest" Then
				temp_type = ""
				temp_amount = ""
				temp_frequency = ""
				temp_housing_payment = ""
				temp_heat_ac = ""
				temp_electric = ""
				temp_phone = ""
				temp_sub_yn = ""
				temp_sub_amt = ""
			ElseIf FORM_QUESTION_ARRAY(form_question_numb).detail_source = "expense" Then
				temp_resident_name = ""
				temp_amount = ""
				temp_current = ""
			ElseIf FORM_QUESTION_ARRAY(form_question_numb).detail_source = "winnings" Then
				temp_resident_name = ""
				temp_amount = ""
				temp_date = ""
			ElseIf FORM_QUESTION_ARRAY(form_question_numb).detail_source = "changes" Then
				temp_resident_name = ""
				temp_date = ""
				temp_explain = ""
			End If
		End If
	Loop until ButtonPressed = return_btn
end function


function save_form_details(FORM_QUESTION_ARRAY)

	xmlPath = user_myDocs_folder & "interview_questions_" & MAXIS_case_number & ".xml"
	Set xmlDoc = nothing
	Set xmlDoc = CreateObject("Microsoft.XMLDOM")
	xmlDoc.async = False

	Set root = xmlDoc.createElement("form")
	xmlDoc.appendChild root

	Set element = xmlDoc.createElement("DHSNumber")
	root.appendChild element
	Set info = xmlDoc.createTextNode(form_number)
	element.appendChild info

	Set element = xmlDoc.createElement("Name")
	root.appendChild element
	Set info = xmlDoc.createTextNode(CAF_form)
	element.appendChild info

	Set element = xmlDoc.createElement("numbOfQuestions")
	root.appendChild element
	Set info = xmlDoc.createTextNode(numb_of_quest)
	element.appendChild info

	Set element = xmlDoc.createElement("lastPageOfQuestions")
	root.appendChild element
	Set info = xmlDoc.createTextNode(last_page_of_questions)
	element.appendChild info

	xmlDoc.save(xmlPath)

	For quest_item = 0 to UBound(FORM_QUESTION_ARRAY)
		' MsgBox "one" & vbCr & quest_item
		Call FORM_QUESTION_ARRAY(quest_item).save_answer(root)
		xmlDoc.save(xmlPath)

	Next

	xmlDoc.save(xmlPath)

	Set xml = CreateObject("Msxml2.DOMDocument")
	Set xsl = CreateObject("Msxml2.DOMDocument")

	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	txt = Replace(fso.OpenTextFile(xmlPath).ReadAll, "><", ">" & vbCrLf & "<")
	stylesheet = "<xsl:stylesheet version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"">" & _
	"<xsl:output method=""xml"" indent=""yes""/>" & _
	"<xsl:template match=""/"">" & _
	"<xsl:copy-of select="".""/>" & _
	"</xsl:template>" & _
	"</xsl:stylesheet>"

	xsl.loadXML stylesheet
	xml.loadXML txt

	xml.transformNode xsl

	xml.Save xmlPath
end function

'constants for TEMP_INFO_ARRAY
const form_yn_const			= 0
const form_second_yn_const	= 1
const form_write_in_const	= 2
const intv_notes_const 		= 3
const verif_yn_const 		= 4
const verif_notes_const		= 5
const q_last_const			= 10

Dim TEMP_INFO_ARRAY()

'CREATE A SUPPLEMENTAL QUESTION PROCESS TO ADD TO QUESTIONS.
const supl_jobs_in_app_month	= 1
const supl_busi_in_app_month	= 2
const supl_unea_in_app_month	= 3
const supl_ac_clarity			= 4
const supl_curr_acct_bal		= 5
const supl_curr_cash_bal		= 6
'THESE can be added to a class as a part of an array so that multiple supplemental questions can be added to a question. These supplemental questions will then display in the dialog.

Dim unea_1_yn, unea_1_amt, unea_2_yn, unea_2_amt, unea_3_yn, unea_3_amt, unea_4_yn, unea_4_amt, unea_5_yn, unea_5_amt
Dim unea_6_yn, unea_6_amt, unea_7_yn, unea_7_amt, unea_8_yn, unea_8_amt, unea_9_yn, unea_9_amt

Dim TEMP_HOUSING_ARRAY()
Dim TEMP_UTILITIES_ARRAY()
ReDim TEMP_HOUSING_ARRAY(6)
ReDim TEMP_UTILITIES_ARRAY(6)
DIM TEMP_ASSETS_ARRAY(4)
DIM TEMP_MSA_ARRAY(3)
DIM TEMP_STWK_ARRAY(3)


question_num = 0
Dim FORM_QUESTION_ARRAY()
ReDim FORM_QUESTION_ARRAY(0)

dim last_page_of_questions, numb_of_quest

If vars_filled = False Then
	If CAF_form = "CAF (DHS-5223)" or CAF_form = "SNAP App for Srs (DHS-5223F)" Then ReDim TEMP_HOUSING_ARRAY(5)
	If CAF_form = "CAF (DHS-5223)" or CAF_form = "SNAP App for Srs (DHS-5223F)" Then ReDim TEMP_UTILITIES_ARRAY(5)

	numb_of_quest = 0
	last_page_of_questions = 4
	Select Case CAF_form
		Case "CAF (DHS-5223)"

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "CAF (DHS-5223)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 1
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does everyone in your household buy, fix or eat food with you?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q1. Does everyone buy, fix, or eat food together?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 1. Does everyone in your household buy, fix or eat food with you?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q1 Information (P&P Together)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 1
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "CAF (DHS-5223)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 2
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Is anyone in the household, who is age 60 or over or disabled, unable to buy or fix food due to a disability?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q2. Is anyone (60+) disabled or unable to prepare food?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 2. Is anyone in the household, who is age 60 or over or disabled, unable to buy or fix food due to a disability?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q2 Information (Ages/DISA unable to buy food)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 2
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "CAF (DHS-5223)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 3
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Is anyone in the household attending school?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q3. Is anyone attending school?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 3. Is anyone in the household attending school?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q3 Information (Attending School)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= True
			FORM_QUESTION_ARRAY(question_num).error_info			= "school"
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= "3. Is anyone in the household attending school? Interview Notes:##~##   - Additional detail about school is needed since this household has children. Gather information about child(ren)'s grade level, district/school, and status.'##~##"
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 3
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "CAF (DHS-5223)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 4
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Is anyone in your household temporarily not living in your home? (eg. vacation, foster care, treatment, hospital, job search)"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q4. Is anyone temporarily not living in the home?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 4. Is anyone in your household temporarily not living in your home? (for example: vacation, foster care, treatment, hospital, job search)"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q4 Information (Temp out of Home)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 4
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "CAF (DHS-5223)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 5
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Is anyone blind, or does anyone have a physical or mental health condition that limits the ability to work or perform daily activities?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q5. Is anyone blind or does anyone have a limiting illness or disability?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 5. Is anyone blind, or does anyone have a physical or mental health condition that limits the ability to work or perform daily activities?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q5 Information (DISA)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 5
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "CAF (DHS-5223)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 6
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Is anyone unable to work for reasons other than illness or disability?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q6. Is anyone unable to work?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 6. Is anyone unable to work for reasons other than illness or disability?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q6 Information (Unable to Work)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 6
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "CAF (DHS-5223)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 7
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "For children under the age of 19, are both parents living in the home?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q7.Are both parents of children under 19 living in the home?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 7. For children under the age of 19, are both parents living in the home?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q7 Information (Both Parents in Home)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 3
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "CAF (DHS-5223)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 8
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "In the last 60 days did anyone in the household:"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q8. In the last 60 days did anyone in the household:"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 8. In the last 60 days did anyone in the household:"
			FORM_QUESTION_ARRAY(question_num).info_type				= "stwk"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= true
			FORM_QUESTION_ARRAY(question_num).make_array_checkboxes = true
			FORM_QUESTION_ARRAY(question_num).item_info_list 		= array("Stop Working or Quit?", "Refuse a Job?", "Ask to Work Fewer Hours?", "Go on Strike?")
			FORM_QUESTION_ARRAY(question_num).item_note_info_list	= array("Stop Working", "Refuse a Job", "Request Fewer Hours", "Strike")
			FORM_QUESTION_ARRAY(question_num).item_ans_list			= array("", "", "", "")
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q8 Information (Job end/reduce in past 60 Days)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 5
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 1
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 100
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "CAF (DHS-5223)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 9
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Has anyone in the household had a job or been self-employed in the past 12 months?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q9. Has anyone had a job OR been self-employed in the past 12 months?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 9. Has anyone in the household had a job or been self-employed in the past 12 months?"
			FORM_QUESTION_ARRAY(question_num).sub_number			= "a"
			FORM_QUESTION_ARRAY(question_num).sub_phrase			= "FOR SNAP ONLY: Has anyone in the household had a job or been self-employed in the past 36 months?"
			FORM_QUESTION_ARRAY(question_num).sub_note_phrase 		= "In the past 36 months? (SNAP ONLY)"
			FORM_QUESTION_ARRAY(question_num).info_type				= "two-part"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q9 Information (Employed in past 12 Months)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 5
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 2
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 75
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "CAF (DHS-5223)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 10
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does anyone in the household have a job or expect to get income from a job this month or next month?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q10. Does anyone have a job?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 10. Does anyone in the household have a job or expect to get income from a job this month or next month?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "jobs"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q10 Information (Job)"
			FORM_QUESTION_ARRAY(question_num).add_to_array_btn		= 3000+question_num

			FORM_QUESTION_ARRAY(question_num).detail_array_exists	= True
			FORM_QUESTION_ARRAY(question_num).detail_source			= "jobs"
			FORM_QUESTION_ARRAY(question_num).detail_button_label 	= "ADD JOB"
			FORM_QUESTION_ARRAY(question_num).detail_interview_notes= array("")
			FORM_QUESTION_ARRAY(question_num).detail_write_in_info	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_verif_status	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_verif_notes	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_edit_btn		= array(2000+question_num*10)
			FORM_QUESTION_ARRAY(question_num).detail_resident_name	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_hourly_wage	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_hours_per_week	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_business		= array("")
			FORM_QUESTION_ARRAY(question_num).detail_monthly_amount	= array("")

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 5
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 3
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 40
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "CAF (DHS-5223)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 11
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Is anyone in the household self-employed or does anyone expect to get income from self-employment this month or next month"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q11.Is anyone self-employed?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 11. Is anyone in the household self-employed or does anyone expect to get income from self-employment this month or next month?"
			FORM_QUESTION_ARRAY(question_num).sub_phrase			= "Gross Earnings"
			FORM_QUESTION_ARRAY(question_num).sub_note_phrase 		= "Gross Monthly Earnings"
			FORM_QUESTION_ARRAY(question_num).info_type				= "single-detail"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q11 Information (Self Employed)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 5
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 4
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "CAF (DHS-5223)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 12
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Do you expect any changes in income, expenses or work hours?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q12.Do you expect any changes in income, expenses, or work hours?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 12. Do you expect any changes in income, expenses or work hours?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q12 Information (Income Changes)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 5
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 5
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "CAF (DHS-5223)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 13
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Has anyone in the household applied for or does anyone get any of the following type of income each month?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q13.Does anyone have any unearned income?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 13. Has anyone in the household applied for or does anyone get any of the following types of income each month?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "unea"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= true
			FORM_QUESTION_ARRAY(question_num).allow_prefil 			= True
			FORM_QUESTION_ARRAY(question_num).item_info_list 		= array("RSDI", "SSI", "VA", "UI", "WC", "Retirement Ben", "Tribal Payments", "CSES", "Other Unearned")
			FORM_QUESTION_ARRAY(question_num).item_note_info_list	= array("RSDI", "SSI", "Veteran Benefits (VA)", "Unemployment Insurance", "Workers' Compensation", "Retirement Benefits", "Tribal payments", "Child or Spousal support", "Other unearned income")
			FORM_QUESTION_ARRAY(question_num).item_ans_list			= array("", "", "", "", "", "", "", "", "")
			FORM_QUESTION_ARRAY(question_num).item_detail_list		= array("", "", "", "", "", "", "", "", "")
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q13 Information (UNEA Income)"
			FORM_QUESTION_ARRAY(question_num).prefil_btn			= 2000+question_num

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 6
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 1
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 130
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "CAF (DHS-5223)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 14
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does anyone in the household have or expect to get any loans, scholarships or grants for attending school?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q14.Does anyone receive financial aid for attending school?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 14. Does anyone in the household have or expect to get any loans, scholarships or grants for attending school?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q14 Information (School Financial Aid)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 6
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 2
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "CAF (DHS-5223)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 15
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does your household have the following housing expenses?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q15.Are there any of the following housing expenses?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 15. Does your household have the following housing expenses?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "housing"

			FORM_QUESTION_ARRAY(question_num).sub_number			= "a"
			FORM_QUESTION_ARRAY(question_num).sub_phrase			= "Do you receive a rental subsidy (ex: Section 8)?"
			FORM_QUESTION_ARRAY(question_num).sub_note_phrase 		= "Do you receive a rental subsidy (ex: Section 8)?"

			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= true
			FORM_QUESTION_ARRAY(question_num).allow_prefil 			= True
			FORM_QUESTION_ARRAY(question_num).item_info_list 		= array("Rent", "Mortgage/contract for deed payment", "Association fees", "Homeowner's insurance", "Room and/or Board", "Real estate taxes")
			FORM_QUESTION_ARRAY(question_num).item_note_info_list	= array("Rent (include mobile home lot rental)", "Mortgage/contract for deed payment", "Association fees", "Homeowner's insurance (if not included in mortgage) ", "Room and/or board", "Real estate taxes (if not included in mortgage)")
			FORM_QUESTION_ARRAY(question_num).item_ans_list			= array("", "", "", "", "", "")
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q15 Information (Housing Expense)"
			FORM_QUESTION_ARRAY(question_num).prefil_btn			= 2000+question_num

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 7
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 1
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 135
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "CAF (DHS-5223)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 16
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does your household have the following utility expenses any time during the year?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q16.Are there any of the following utility expenses?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 16. Does your household have the following utility expenses any time during the year?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "utilities"

			FORM_QUESTION_ARRAY(question_num).sub_number			= "a"
			FORM_QUESTION_ARRAY(question_num).sub_phrase			= "Did you or anyone in your household receive energy assistance of more than $20 in the past 12 months?"
			FORM_QUESTION_ARRAY(question_num).sub_note_phrase 		= "Did anyone receive energy assistance?"

			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= true
			FORM_QUESTION_ARRAY(question_num).item_info_list 		= array("Heating/air conditioning", "Electricity", "Cooking fuel", "Water and sewer", "Garbage removal", "Phone/cell phone")
			FORM_QUESTION_ARRAY(question_num).item_note_info_list	= array("Heat/AC", "Electric", "Cooking Fuel", "Water/Sewer", "Garbage", "Phone")
			FORM_QUESTION_ARRAY(question_num).item_ans_list			= array("", "", "", "", "", "")
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q16 Information (Utilities Expense)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 7
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 2
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 120
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "CAF (DHS-5223)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 17
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Do you or anyone living with you have costs for care of a child(ren) because you or they are working, looking for work or going to school?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q17.Does anyone have costs for childcare?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 17. Do you or anyone living with you have costs for care of a child(ren) because you or they are working, looking for work or going to school?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q17 Information (Child Care Expense)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 8
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 1
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "CAF (DHS-5223)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 18
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does anyone have costs for care of an ill/disabled adult because you or they are working, looking for work or going to school?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q18.Does anyone have costs for adult care?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 18. Do you or anyone living with you have costs for care of an ill or disabled adult because you or they are working, looking for work or going to school?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q18 Information (DISA Adult Care Expense)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 8
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 2
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "CAF (DHS-5223)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 19
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does anyone in the household pay support, or contribute to a tax dependent who does not live in your home?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q19.Does anyone pay support to someone outside of the home?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 19. Does anyone in the household pay court-ordered child support, spousal support, child care support, medical support or contribute to a tax dependent who does not live in your home?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q19 Information (Child Support Expense)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 8
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 3
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "CAF (DHS-5223)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 20
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "For SNAP only: Does anyone in the household have medical expenses?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q20.Does anyone (disabled or 60+) have medical expenses? (SNAP ONLY)"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 20. For SNAP only: Does anyone in the household have medical expenses? "
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q20 Information (Medical Expenses)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 8
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 4
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "CAF (DHS-5223)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 21
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does anyone in the household own, or is anyone buying, any of the following?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q21.Does anyone own or is anyone buying any of the following:"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 21. Does anyone in the household own, or is anyone buying, any of the following? Check yes or no for each item. "
			FORM_QUESTION_ARRAY(question_num).info_type				= "assets"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= true
			FORM_QUESTION_ARRAY(question_num).item_info_list 		= array("Cash", "Bank accounts (savings, checking, etc)", "Stocks, bonds, annuities, 401k, etc", "Vehicles (cars, trucks, motorcycles, campers, trailers)", "Electronic Payment Card (Reliacard, debit, etc.)")
			FORM_QUESTION_ARRAY(question_num).item_note_info_list	= array("Cash", "Bank Accounts", "Stocks", "Vehicles", "Payment Card")
			FORM_QUESTION_ARRAY(question_num).item_ans_list			= array("", "", "", "", "")
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q21 Information (Assets)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 8
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 5
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 105
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "CAF (DHS-5223)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 22
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "For Cash programs only: Has anyone in the household given away, sold or traded anything of value in the past 12 months?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q22.Has anyone sold/given away/traded assets in the past 12 mos?(CASH ONLY)"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 22. For Cash programs only: Has anyone in the household given away, sold or traded anything of value in the past 12 months? (For example: Cash, Bank accounts, Stocks, Bonds, Vehicles)"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q22 Information (Asset Trade)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 9
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 1
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "CAF (DHS-5223)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 23
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "For recertifications only: Did anyone move in or out of your home in the past 12 months?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q23.Did anyone move in/out in the past 12 months? (REVW ONLY)"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 23. For recertifications only: Did anyone move in or out of your home in the past 12 months?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q23 Information (Anyone Move In or Out)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 9
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 2
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "CAF (DHS-5223)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 24
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "For MSA recipients only: Does anyone in the household have any of the following expenses?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q24.Does anyone have any of the following expenses? (MSA ONLY)"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 24. For MSA recipients only: Does anyone in the household have any of the following expenses?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "msa"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= true
			FORM_QUESTION_ARRAY(question_num).item_info_list 		= array("Representative Payee fees", "Guardian Conservator fees", "Physician-perscribed special diet", "High housing costs")
			FORM_QUESTION_ARRAY(question_num).item_note_info_list	= array("REP Payee Fees", "Guard Fees", "Special Diet", "High Housing Costs")
			FORM_QUESTION_ARRAY(question_num).item_ans_list			= array("", "", "", "")
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q24 Information (MSA Expenses)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 9
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 4
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 105
			question_num = question_num + 1

			numb_of_quest = question_num-1
			last_page_of_questions = 9


		Case "MNbenefits"

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "MNbenefits"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 1
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does everyone in your household buy, fix or eat food with you?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q1. Does everyone buy, fix, or eat food together?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 1. Does everyone in your household buy, fix or eat food with you?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q1 Information (P&P Together)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 1
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "MNbenefits"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 2
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Is anyone in the household, who is age 60 or over or disabled, unable to buy or fix food due to a disability?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q2. Is anyone (60+) disabled or unable to prepare food?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 2. Is anyone in the household, who is age 60 or over or disabled, unable to buy or fix food due to a disability?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q2 Information (Ages/DISA unable to buy food)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 2
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "MNbenefits"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 3
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Is anyone in the household attending school?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q3. Is anyone attending school?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 3. Is anyone in the household attending school?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q3 Information (Attending School)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 3
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "MNbenefits"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 4
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Is anyone in your household temporarily not living in your home? (eg. vacation, foster care, treatment, hospital, job search)"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q4. Is anyone temporarily not living in the home?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 4. Is anyone in your household temporarily not living in your home? (for example: vacation, foster care, treatment, hospital, job search)"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q4 Information (Temp out of Home)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 4
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "MNbenefits"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 5
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Is anyone blind, or does anyone have a physical or mental health condition that limits the ability to work or perform daily activities?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q5. Is anyone blind or does anyone have a limiting illness or disability?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 5. Is anyone blind, or does anyone have a physical or mental health condition that limits the ability to work or perform daily activities?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q5 Information (DISA)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 5
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "MNbenefits"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 6
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Is anyone unable to work for reasons other than illness or disability?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q6. Is anyone unable to work?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 6. Is anyone unable to work for reasons other than illness or disability?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q6 Information (Unable to Work)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 6
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "MNbenefits"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 7
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "In the last 60 days did anyone in the household: - Stop working or quit a job? - Refuse a job offer? - Ask to work fewer hours? - Go on strike?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q7. Has anyone stopped, quit or refused employment in the past 60 days?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 7. In the last 60 days did anyone in the household: - Stop working or quit a job? - Refuse a job offer? - Ask to work fewer hours? - Go on strike?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q7 Information (Job end/reduce in past 60 Days)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 5
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 1
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "MNbenefits"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 8
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Has anyone in the household had a job or been self-employed in the past 12 months?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q8. Has anyone had a job OR been self-employed in the past 12 months?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 8. Has anyone in the household had a job or been self-employed in the past 12 months?"
			FORM_QUESTION_ARRAY(question_num).sub_number			= "a"
			FORM_QUESTION_ARRAY(question_num).sub_phrase			= "FOR SNAP ONLY: Has anyone in the household had a job or been self-employed in the past 36 months?"
			FORM_QUESTION_ARRAY(question_num).sub_note_phrase 		= "In the past 36 months? (SNAP ONLY)"
			FORM_QUESTION_ARRAY(question_num).info_type				= "two-part"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q8 Information (Employed in past 12 Months)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 5
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 2
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 75
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "MNbenefits"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 9
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does anyone in the household have a job or expect to get income from a job this month or next month?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q9. Does anyone have a job?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 9. Does anyone in the household have a job or expect to get income from a job this month or next month?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "jobs"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q9 Information (Job)"
			FORM_QUESTION_ARRAY(question_num).add_to_array_btn		= 3000+question_num

			FORM_QUESTION_ARRAY(question_num).detail_array_exists	= True
			FORM_QUESTION_ARRAY(question_num).detail_source			= "jobs"
			FORM_QUESTION_ARRAY(question_num).detail_button_label 	= "ADD JOB"
			FORM_QUESTION_ARRAY(question_num).detail_interview_notes= array("")
			FORM_QUESTION_ARRAY(question_num).detail_write_in_info	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_verif_status	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_verif_notes	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_edit_btn		= array(2000+question_num*10)
			FORM_QUESTION_ARRAY(question_num).detail_resident_name	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_hourly_wage	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_hours_per_week	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_business		= array("")
			FORM_QUESTION_ARRAY(question_num).detail_monthly_amount	= array("")

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 5
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 3
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 40
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "MNbenefits"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 10
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Is anyone in the household self-employed or does anyone expect to get income from self-employment this month or next month"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q10.Is anyone self-employed?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 10. Is anyone in the household self-employed or does anyone expect to get income from self-employment this month or next month?"
			FORM_QUESTION_ARRAY(question_num).sub_phrase			= "Gross Earnings"
			FORM_QUESTION_ARRAY(question_num).sub_note_phrase 		= "Gross Monthly Earnings"
			FORM_QUESTION_ARRAY(question_num).info_type				= "single-detail"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q10 Information (Self Employed)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 5
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 4
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "MNbenefits"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 11
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Do you expect any changes in income, expenses or work hours?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q11.Do you expect any changes in income, expenses, or work hours?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 11. Do you expect any changes in income, expenses or work hours?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q11 Information (Income Changes)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 5
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 5
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "MNbenefits"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 12
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Has anyone in the household applied for or does anyone get any of the following type of income each month?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q12.Does anyone have any unearned income?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 12. Has anyone in the household applied for or does anyone get any of the following types of income each month?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "unea"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= true
			FORM_QUESTION_ARRAY(question_num).allow_prefil 			= True
			FORM_QUESTION_ARRAY(question_num).item_info_list 		= array("RSDI", "SSI", "VA", "UI", "WC", "Retirement Ben", "Tribal Payments", "CSES", "Other Unearned")
			FORM_QUESTION_ARRAY(question_num).item_note_info_list	= array("RSDI", "SSI", "Veteran Benefits (VA)", "Unemployment Insurance", "Workers' Compensation", "Retirement Benefits", "Tribal payments", "Child or Spousal support", "Other unearned income")
			FORM_QUESTION_ARRAY(question_num).item_ans_list			= array("", "", "", "", "", "", "", "", "")
			FORM_QUESTION_ARRAY(question_num).item_detail_list		= array("", "", "", "", "", "", "", "", "")
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q12 Information (UNEA Income)"
			FORM_QUESTION_ARRAY(question_num).prefil_btn			= 2000+question_num

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 6
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 1
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 130
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "MNbenefits"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 13
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does anyone in the household have or expect to get any loans, scholarships or grants for attending school?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q13.Does anyone receive financial aid for attending school?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 13. Does anyone in the household have or expect to get any loans, scholarships or grants for attending school?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q13 Information (School Financial Aid)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 6
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 2
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "MNbenefits"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 14
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does your household have the following housing expenses?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q14.Are there any of the following housing expenses?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 14. Does your household have the following housing expenses?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "housing"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= true
			FORM_QUESTION_ARRAY(question_num).allow_prefil 			= True
			FORM_QUESTION_ARRAY(question_num).item_info_list 		= array("Rent", "Rental or Section 8 Subsidy", "Mortgage/contract for deed payment", "Association fees", "Homeowner's insurance", "Room and/or Board", "Real estate taxes")
			FORM_QUESTION_ARRAY(question_num).item_note_info_list	= array("Rent (include mobile home lot rental)", "Rent or Section 8 subsidy", "Mortgage/contract for deed payment", "Association fees", "Homeowner's insurance (if not included in mortgage) ", "Room and/or board", "Real estate taxes (if not included in mortgage)")
			FORM_QUESTION_ARRAY(question_num).item_ans_list			= array("", "", "", "", "", "", "")
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q14 Information (Housing Expense)"
			FORM_QUESTION_ARRAY(question_num).prefil_btn			= 2000+question_num

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 7
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 1
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 135
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "MNbenefits"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 15
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does your household have the following utility expenses any time during the year?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q15.Are there any of the following utility expenses?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 15. Does your household have the following utility expenses any time during the year?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "utilities"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= true
			FORM_QUESTION_ARRAY(question_num).item_info_list 		= array("Heating/air conditioning", "Electricity", "Cooking fuel", "Water and sewer", "Garbage removal", "Phone/cell phone", "Did you or anyone in your houehold receive LIHEAP (energy assistance) for more than $20 in the past 12 months?")
			FORM_QUESTION_ARRAY(question_num).item_note_info_list	= array("Heat/AC", "Electric", "Cooking Fuel", "Water/Sewer", "Garbage", "Phone", "LIHEAP/Energy Assistance in past 12 months")
			FORM_QUESTION_ARRAY(question_num).item_ans_list			= array("", "", "", "", "", "", "")
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q15 Information (Utilities Expense)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 7
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 2
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 120
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "MNbenefits"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 16
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Do you or anyone living with you have costs for care of a child(ren) because you or they are working, looking for work or going to school?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q16.Does anyone have costs for childcare?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 16. Do you or anyone living with you have costs for care of a child(ren) because you or they are working, looking for work or going to school?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q16 Information (Child Care Expense)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 8
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 1
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "MNbenefits"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 17
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does anyone have costs for care of an ill/disabled adult because you or they are working, looking for work or going to school?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q17.Does anyone have costs for adult care?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 17. Do you or anyone living with you have costs for care of an ill or disabled adult because you or they are working, looking for work or going to school?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q17 Information (DISA Adult Care Expense)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 8
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 2
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "MNbenefits"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 18
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does anyone in the household pay support, or contribute to a tax dependent who does not live in your home?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q18.Does anyone pay support to someone outside of the home?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 18. Does anyone in the household pay court-ordered child support, spousal support, child care support, medical support or contribute to a tax dependent who does not live in your home?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q18 Information (Child Support Expense)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 8
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 3
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "MNbenefits"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 19
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "For SNAP only: Does anyone in the household have medical expenses?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q19.Does anyone (disabled or 60+) have medical expenses? (SNAP ONLY)"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 19. For SNAP only: Does anyone in the household have medical expenses? "
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q19 Information (Medical Expenses)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 8
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 4
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "MNbenefits"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 20
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does anyone in the household own, or is anyone buying, any of the following?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q20.Does anyone own or is anyone buying any of the following:"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 20. Does anyone in the household own, or is anyone buying, any of the following? Check yes or no for each item. "
			FORM_QUESTION_ARRAY(question_num).info_type				= "assets"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= true
			FORM_QUESTION_ARRAY(question_num).item_info_list 		= array("Cash", "Bank accounts (savings, checking, etc)", "Stocks, bonds, annuities, 401k, etc", "Vehicles (cars, trucks, motorcycles, campers, trailers)", "Electronic Payment Card (Reliacard, debit, etc.)")
			FORM_QUESTION_ARRAY(question_num).item_note_info_list	= array("Cash", "Bank Accounts", "Stocks", "Vehicles", "Payment Card")
			FORM_QUESTION_ARRAY(question_num).item_ans_list			= array("", "", "", "", "")
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q20 Information (Assets)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 8
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 5
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 105
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "MNbenefits"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 21
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "For Cash programs only: Has anyone in the household given away, sold or traded anything of value in the past 12 months?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q21.Has anyone sold/given away/traded assets in the past 12 mos?(CASH ONLY)"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 21. For Cash programs only: Has anyone in the household given away, sold or traded anything of value in the past 12 months? (For example: Cash, Bank accounts, Stocks, Bonds, Vehicles)"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q21 Information (Asset Trade)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 9
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 1
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "MNbenefits"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 22
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "For recertifications only: Did anyone move in or out of your home in the past 12 months?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q22.Did anyone move in/out in the past 12 months? (REVW ONLY)"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 22. For recertifications only: Did anyone move in or out of your home in the past 12 months?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q22 Information (Anyone Move In or Out)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 9
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 2
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "MNbenefits"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 23
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "For children under the age of 19, are both parents living in the home?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q23.Are both parents of children under 19 living in the home?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 23. For children under the age of 19, are both parents living in the home?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q23 Information (Both Parents in Home)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 9
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 3
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "MNbenefits"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 24
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "For MSA recipients only: Does anyone in the household have any of the following expenses?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q24.Does anyone have any of the following expenses? (MSA ONLY)"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 24. For MSA recipients only: Does anyone in the household have any of the following expenses?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "msa"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= true
			FORM_QUESTION_ARRAY(question_num).item_info_list 		= array("Representative Payee fees", "Guardian Conservator fees", "Physician-perscribed special diet", "High housing costs")
			FORM_QUESTION_ARRAY(question_num).item_note_info_list	= array("REP Payee Fees", "Guard Fees", "Special Diet", "High Housing Costs")
			FORM_QUESTION_ARRAY(question_num).item_ans_list			= array("", "", "", "")
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "CAF Q24 Information (MSA Expenses)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 9
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 4
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 105
			question_num = question_num + 1

			numb_of_quest = question_num-1
			last_page_of_questions = 9


		Case "HUF (DHS-8107)"
			' ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "HUF (DHS-8107)"
			' Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			' FORM_QUESTION_ARRAY(question_num).number 				= 1
			' FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "How can we send you updates and reminders about your case in the future?"
			' FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q1. How can we send you updates/ reminders about your case in the future?"
			' FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 1. How can we send you updates/ reminders about your case in the future?"
			' FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			' FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			' FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			' FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			' FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			' FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
			' FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			' question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "HUF (DHS-8107)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 2
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Has anyone moved in or out of your home since your last review or application?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q2. Has anyone moved in or out of the home?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 2. Has anyone moved in or out of the home?"
			FORM_QUESTION_ARRAY(question_num).sub_number			= "a"
			FORM_QUESTION_ARRAY(question_num).sub_phrase			= "Does everyone in your household buy, fix or eat food with you?"
			FORM_QUESTION_ARRAY(question_num).sub_note_phrase 		= "Does everyone buy, fix, or eat food together?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "two-part"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "HUF Q2 Information (Anyone Move In or Out)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 75
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "HUF (DHS-8107)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 3
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does anyone in your household have assets?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q3. Does anyone in your household have assets?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 3. Does anyone in your household have assets?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "HUF Q3 Information (Assets)"
			FORM_QUESTION_ARRAY(question_num).detail_array_exists	= True
			FORM_QUESTION_ARRAY(question_num).detail_source			= "assets"
			FORM_QUESTION_ARRAY(question_num).detail_button_label 	= "ADD ASSET"
			FORM_QUESTION_ARRAY(question_num).detail_interview_notes= array("")
			FORM_QUESTION_ARRAY(question_num).detail_write_in_info	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_verif_status	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_verif_notes	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_edit_btn		= array("")
			FORM_QUESTION_ARRAY(question_num).detail_edit_btn		= array(2000+question_num*10)
			FORM_QUESTION_ARRAY(question_num).detail_resident_name	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_type			= array("")
			FORM_QUESTION_ARRAY(question_num).detail_value			= array("")
			FORM_QUESTION_ARRAY(question_num).detail_explain		= array("")

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "HUF (DHS-8107)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 4
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does anyone in the household have a job or expect to get income from a job this month or next month?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q4. Does anyone have a job?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 4. Does anyone in the household have a job or expect to get income from a job this month or next month?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "HUF Q4 Information (Job)"
			FORM_QUESTION_ARRAY(question_num).add_to_array_btn		= 3000+question_num

			FORM_QUESTION_ARRAY(question_num).detail_array_exists	= True
			FORM_QUESTION_ARRAY(question_num).detail_source			= "jobs"
			FORM_QUESTION_ARRAY(question_num).detail_button_label 	= "ADD JOB"
			FORM_QUESTION_ARRAY(question_num).detail_interview_notes= array("")
			FORM_QUESTION_ARRAY(question_num).detail_write_in_info	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_verif_status	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_verif_notes	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_edit_btn		= array("")
			FORM_QUESTION_ARRAY(question_num).detail_edit_btn		= array(2000+question_num*10)
			FORM_QUESTION_ARRAY(question_num).detail_resident_name	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_hourly_wage	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_hours_per_week	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_business		= array("")
			FORM_QUESTION_ARRAY(question_num).detail_monthly_amount	= array("")

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 5
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "HUF (DHS-8107)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 5
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Is anyone in the household self-employed or does anyone expect to get income from self-employment this month or next month?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q5. Is anyone self-employed?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 5. Is anyone in the household self-employed or does anyone expect to get income from self-employment this month or next month?"
			FORM_QUESTION_ARRAY(question_num).sub_phrase			= "Gross Earnings"
			FORM_QUESTION_ARRAY(question_num).sub_note_phrase 		= "Gross Monthly Earnings"
			FORM_QUESTION_ARRAY(question_num).info_type				= "single-detail"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "HUF Q5 Information (Self Employed)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 5
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "HUF (DHS-8107)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 6
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does anyone in your household get money or expect to get money from sources other than work?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q6. Does anyone have any unearned income?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 6. Does anyone in your household get money or expect to get money from sources other than work?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "HUF Q6 Information (UNEA Income)"
			FORM_QUESTION_ARRAY(question_num).add_to_array_btn		= 3000+question_num

			FORM_QUESTION_ARRAY(question_num).detail_array_exists	= True
			FORM_QUESTION_ARRAY(question_num).detail_source			= "unea"
			FORM_QUESTION_ARRAY(question_num).detail_button_label 	= "ADD INCOME"
			FORM_QUESTION_ARRAY(question_num).detail_interview_notes= array("")
			FORM_QUESTION_ARRAY(question_num).detail_write_in_info	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_verif_status	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_verif_notes	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_edit_btn		= array("")
			FORM_QUESTION_ARRAY(question_num).detail_edit_btn		= array(2000+question_num*10)
			FORM_QUESTION_ARRAY(question_num).detail_resident_name	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_type			= array("")
			FORM_QUESTION_ARRAY(question_num).detail_date			= array("")
			FORM_QUESTION_ARRAY(question_num).detail_amount			= array("")
			FORM_QUESTION_ARRAY(question_num).detail_frequency		= array("")

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 5
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "HUF (DHS-8107)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 7
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Have your shelter and/or utility costs changed since your last review or application?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q7. Have any housing costs changed?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 7. Have any shelter/utility costs changed?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "HUF Q7 Information (Housing Expense)"
			FORM_QUESTION_ARRAY(question_num).add_to_array_btn		= 3000+question_num

			FORM_QUESTION_ARRAY(question_num).detail_array_exists	= True
			FORM_QUESTION_ARRAY(question_num).detail_source			= "shel-hest"
			FORM_QUESTION_ARRAY(question_num).detail_button_label 	= "ADD EXPENSE"
			FORM_QUESTION_ARRAY(question_num).detail_interview_notes= array("")
			FORM_QUESTION_ARRAY(question_num).detail_write_in_info	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_verif_status	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_verif_notes	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_edit_btn		= array("")
			FORM_QUESTION_ARRAY(question_num).detail_edit_btn		= array(2000+question_num*10)
			FORM_QUESTION_ARRAY(question_num).detail_type			= array("")
			FORM_QUESTION_ARRAY(question_num).detail_amount			= array("")
			FORM_QUESTION_ARRAY(question_num).detail_frequency		= array("")

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 6
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1
			'NEED TO ADD MORE QUESTIONS HERE FOR THIS QUESTION

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "HUF (DHS-8107)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 8
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does anyone in your household pay court-ordered child or medical support?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q8. Does anyone pay support to someone outside of the home?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 8. Does anyone in your household pay court-ordered child or medical support?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "HUF Q8 Information (Child Support Expense)"
			FORM_QUESTION_ARRAY(question_num).add_to_array_btn		= 3000+question_num

			FORM_QUESTION_ARRAY(question_num).detail_array_exists	= True
			FORM_QUESTION_ARRAY(question_num).detail_source			= "expense"
			FORM_QUESTION_ARRAY(question_num).detail_button_label 	= "ADD EXPENSE"
			FORM_QUESTION_ARRAY(question_num).detail_interview_notes= array("")
			FORM_QUESTION_ARRAY(question_num).detail_write_in_info	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_verif_status	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_verif_notes	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_edit_btn		= array("")
			FORM_QUESTION_ARRAY(question_num).detail_edit_btn		= array(2000+question_num*10)
			FORM_QUESTION_ARRAY(question_num).detail_resident_name	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_amount			= array("")
			FORM_QUESTION_ARRAY(question_num).detail_current		= array("")

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 6
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "HUF (DHS-8107)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 9
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Do you or anyone living with you have costs for care of a child(ren) or adult due to work, looking for work, or school?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q9. Does anyone have costs for childcare?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 9. Do you or anyone living with you have costs for care of a child(ren) or an ill or disabled adult because you or they are working, looking for work, or going to school?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "HUF Q9 Information (Child Care Expense)"
			FORM_QUESTION_ARRAY(question_num).add_to_array_btn		= 3000+question_num

			FORM_QUESTION_ARRAY(question_num).detail_array_exists	= True
			FORM_QUESTION_ARRAY(question_num).detail_source			= "expense"
			FORM_QUESTION_ARRAY(question_num).detail_button_label 	= "ADD EXPENSE"
			FORM_QUESTION_ARRAY(question_num).detail_interview_notes= array("")
			FORM_QUESTION_ARRAY(question_num).detail_write_in_info	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_verif_status	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_verif_notes	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_edit_btn		= array("")
			FORM_QUESTION_ARRAY(question_num).detail_edit_btn		= array(2000+question_num*10)
			FORM_QUESTION_ARRAY(question_num).detail_resident_name	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_amount			= array("")
			FORM_QUESTION_ARRAY(question_num).detail_current		= array("")

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 6
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "HUF (DHS-8107)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 10
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does anyone in your household have changes to health care expenses, or new expenses?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q10. Any changes to medical expenses?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 10. Does anyone in your household have changes to health care expenses, or new expenses?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "HUF Q10 Information (Medical Expenses)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 6
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "HUF (DHS-8107)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 11
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does anyone in your household have or expect to get any loans, scholarships or grants for attending school?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q11. Does anyone receive financial aid for attending school?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 11. Does anyone in your household have or expect to get any loans, scholarships or grants for attending school?"
			FORM_QUESTION_ARRAY(question_num).sub_phrase			= "Who?"
			FORM_QUESTION_ARRAY(question_num).sub_note_phrase 		= "Who?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "single-detail"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "HUF Q11 Information (School Financial Aid)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 7
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			'THIS IS REALLY A CHECKBOX SITUATION
			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "HUF (DHS-8107)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 12
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "For Minnesota Supplemental Aid (MSA) recipients: do you have any of the following expenses?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q12. Does anyone have any of the following expenses? (MSA ONLY)"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 12. For Minnesota Supplemental Aid (MSA) recipients: do you have any of the following expenses?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "msa"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= true
			FORM_QUESTION_ARRAY(question_num).make_array_checkboxes = true
			FORM_QUESTION_ARRAY(question_num).item_info_list 		= array("Representative Payee fees", "Guardian Conservator fees", "Physician-perscribed special diet", "High housing costs")
			FORM_QUESTION_ARRAY(question_num).item_note_info_list	= array("REP Payee Fees", "Guard Fees", "Special Diet", "High Housing Costs")
			FORM_QUESTION_ARRAY(question_num).item_ans_list			= array("", "", "", "")
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "HUF Q12 Information (MSA Expenses)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 7
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 100
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "HUF (DHS-8107)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 13
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Other changes"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q13. Other changes"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 13. Other changes"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "HUF Q13 Information (Changes)"
			FORM_QUESTION_ARRAY(question_num).add_to_array_btn		= 3000+question_num

			FORM_QUESTION_ARRAY(question_num).detail_array_exists	= True
			FORM_QUESTION_ARRAY(question_num).detail_source			= "changes"
			FORM_QUESTION_ARRAY(question_num).detail_button_label 	= "ADD CHANGE"
			FORM_QUESTION_ARRAY(question_num).detail_interview_notes= array("")
			FORM_QUESTION_ARRAY(question_num).detail_write_in_info	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_verif_status	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_verif_notes	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_edit_btn		= array("")
			FORM_QUESTION_ARRAY(question_num).detail_edit_btn		= array(2000+question_num*10)
			FORM_QUESTION_ARRAY(question_num).detail_resident_name	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_date			= array("")
			FORM_QUESTION_ARRAY(question_num).detail_explain		= array("")

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 7
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			numb_of_quest = question_num-1
			last_page_of_questions = 7




		Case "SNAP App for Srs (DHS-5223F)"

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "SNAP App for Srs (DHS-5223F)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 1
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does anyone in the household have a job or expect to get income from a job this month or next month?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q1. Does anyone have a job?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 1. Does anyone in the household have a job or expect to get income from a job this month or next month?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "jobs"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "SR SNAP APP Q1 Information (Job)"
			FORM_QUESTION_ARRAY(question_num).add_to_array_btn		= 3000+question_num

			FORM_QUESTION_ARRAY(question_num).detail_array_exists	= True
			FORM_QUESTION_ARRAY(question_num).detail_source			= "jobs"
			FORM_QUESTION_ARRAY(question_num).detail_button_label 	= "ADD JOB"
			FORM_QUESTION_ARRAY(question_num).detail_interview_notes= array("")
			FORM_QUESTION_ARRAY(question_num).detail_write_in_info	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_verif_status	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_verif_notes	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_edit_btn		= array(2000+question_num*10)
			FORM_QUESTION_ARRAY(question_num).detail_resident_name	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_hourly_wage	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_hours_per_week	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_business		= array("")
			FORM_QUESTION_ARRAY(question_num).detail_monthly_amount	= array("")

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 3
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 40
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "SNAP App for Srs (DHS-5223F)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 2
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Is anyone in the household self-employed or does anyone expect to get income from self-employment this month or next month"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q2.Is anyone self-employed?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 2. Is anyone in the household self-employed or does anyone expect to get income from self-employment this month or next month?"
			FORM_QUESTION_ARRAY(question_num).sub_phrase			= "Gross Earnings"
			FORM_QUESTION_ARRAY(question_num).sub_note_phrase 		= "Gross Monthly Earnings"
			FORM_QUESTION_ARRAY(question_num).info_type				= "single-detail"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "SR SNAP APP Q2 Information (Self Employed)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 4
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "SNAP App for Srs (DHS-5223F)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 3
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Has anyone in the household applied for or does anyone get any of the following type of income each month?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q3.Does anyone have any unearned income?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 3. Has anyone in the household applied for or does anyone get any of the following types of income each month?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "unea"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= true
			FORM_QUESTION_ARRAY(question_num).allow_prefil 			= True
			FORM_QUESTION_ARRAY(question_num).item_info_list 		= array("RSDI", "SSI", "VA", "UI", "WC", "Retirement Ben", "Tribal Payments", "CSES", "Other Unearned")
			FORM_QUESTION_ARRAY(question_num).item_note_info_list	= array("RSDI", "SSI", "Veteran Benefits (VA)", "Unemployment Insurance", "Workers' Compensation", "Retirement Benefits", "Tribal payments", "Child or Spousal support", "Other unearned income")
			FORM_QUESTION_ARRAY(question_num).item_ans_list			= array("", "", "", "", "", "", "", "", "")
			FORM_QUESTION_ARRAY(question_num).item_detail_list		= array("", "", "", "", "", "", "", "", "")
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "SR SNAP APP Q3 Information (UNEA Income)"
			FORM_QUESTION_ARRAY(question_num).prefil_btn			= 2000+question_num

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 1
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 130
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "SNAP App for Srs (DHS-5223F)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 4
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does your household have the following housing expenses?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q4.Are there any of the following housing expenses?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 4. Does your household have the following housing expenses?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "housing"
			FORM_QUESTION_ARRAY(question_num).sub_number			= "a"
			FORM_QUESTION_ARRAY(question_num).sub_phrase			= "Do you receive a rental subsidy (ex: Section 8)?"
			FORM_QUESTION_ARRAY(question_num).sub_note_phrase 		= "Do you receive a rental subsidy (ex: Section 8)?"

			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= true
			FORM_QUESTION_ARRAY(question_num).allow_prefil 			= True
			FORM_QUESTION_ARRAY(question_num).item_info_list 		= array("Rent", "Mortgage/contract for deed payment", "Association fees", "Homeowner's insurance", "Room and/or Board", "Real estate taxes")
			FORM_QUESTION_ARRAY(question_num).item_note_info_list	= array("Rent (include mobile home lot rental)", "Mortgage/contract for deed payment", "Association fees", "Homeowner's insurance (if not included in mortgage) ", "Room and/or board", "Real estate taxes (if not included in mortgage)")
			FORM_QUESTION_ARRAY(question_num).item_ans_list			= array("", "", "", "", "", "")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "SR SNAP APP Q4 Information (Housing Expense)"
			FORM_QUESTION_ARRAY(question_num).prefil_btn			= 2000+question_num

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 5
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 1
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 135
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "SNAP App for Srs (DHS-5223F)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 5
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does your household have the following utility expenses any time during the year?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q5.Are there any of the following utility expenses?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 5. Does your household have the following utility expenses any time during the year?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "utilities"
			FORM_QUESTION_ARRAY(question_num).sub_number			= "a"
			FORM_QUESTION_ARRAY(question_num).sub_phrase			= "Did you or anyone in your household receive energy assistance of more than $20 in the past 12 months?"
			FORM_QUESTION_ARRAY(question_num).sub_note_phrase 		= "Did anyone receive energy assistance?"

			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= true
			FORM_QUESTION_ARRAY(question_num).item_info_list 		= array("Heating/air conditioning", "Electricity", "Cooking fuel", "Water and sewer", "Garbage removal", "Phone/cell phone")
			FORM_QUESTION_ARRAY(question_num).item_note_info_list	= array("Heat/AC", "Electric", "Cooking Fuel", "Water/Sewer", "Garbage", "Phone")
			FORM_QUESTION_ARRAY(question_num).item_ans_list			= array("", "", "", "", "", "")
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "SR SNAP APP Q5 Information (Utilities Expense)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 5
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 2
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 120
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "SNAP App for Srs (DHS-5223F)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 6
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does anyone have costs for care of an ill/disabled adult because you or they are working, looking for work or going to school?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q6.Does anyone have costs for adult care?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 6. Do you or anyone living with you have costs for care of an ill or disabled adult because you or they are working, looking for work or going to school?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "SR SNAP APP Q6 Information (DISA Adult Care Expense)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 5
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 2
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "SNAP App for Srs (DHS-5223F)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 7
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does anyone in the household pay support, or contribute to a tax dependent who does not live in your home?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q7.Does anyone pay support to someone outside of the home?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 7. Does anyone in the household pay court-ordered child support, spousal support, child care support, medical support or contribute to a tax dependent who does not live in your home?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "SR SNAP APP Q7 Information (Support Expense)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 6
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 3
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "SNAP App for Srs (DHS-5223F)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 8
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "For SNAP only: Does anyone in the household have medical expenses?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q8.Does anyone (disabled or 60+) have medical expenses? (SNAP ONLY)"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 8. For SNAP only: Does anyone in the household have medical expenses? "
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "SR SNAP APP Q8 Information (Medical Expenses)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 6
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 4
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			numb_of_quest = question_num-1
			last_page_of_questions = 6

		Case "Combined AR for Certain Pops (DHS-3727)"


			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "Combined AR for Certain Pops (DHS-3727)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 2
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Do you or your spouse have any changes from the last year?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q2. Any changes in the last year?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 2. Do you or your spouse have any changes from the last year?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "Combined AR Q2 Information (Changes)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 1
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "Combined AR for Certain Pops (DHS-3727)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 3
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Do you or your spouse have any assets that require proof?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q3. Does anyone in your household have assets?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 3. Do you or your spouse have any assets that require proof?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "Combined AR Q3 Information (Assets)"
			FORM_QUESTION_ARRAY(question_num).detail_array_exists	= True
			FORM_QUESTION_ARRAY(question_num).detail_source			= "assets"
			FORM_QUESTION_ARRAY(question_num).detail_button_label 	= "ADD ASSET"
			FORM_QUESTION_ARRAY(question_num).detail_interview_notes= array("")
			FORM_QUESTION_ARRAY(question_num).detail_write_in_info	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_verif_status	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_verif_notes	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_edit_btn		= array("")
			FORM_QUESTION_ARRAY(question_num).detail_edit_btn		= array(2000+question_num*10)
			FORM_QUESTION_ARRAY(question_num).detail_resident_name	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_type			= array("")
			FORM_QUESTION_ARRAY(question_num).detail_value			= array("")
			FORM_QUESTION_ARRAY(question_num).detail_explain		= array("")

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "Combined AR for Certain Pops (DHS-3727)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 4
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "For MA-LTC, did you or your spouse: - Buy, sell, trade, or give away assets - or refuse income or assets? - Purchase an annuity, life estate, promissory note, loan, mortgage, or create a trust?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q4. Did you buy, sell, or trade assets or income? (MA-LTC)"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 4. For MA-LTC, did you buy, sell, or trade assets or income? (For MA-LTC)"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "Combined AR Q4 Information (Asset Trade)"

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
			FORM_QUESTION_ARRAY(question_num).dialog_order 			= 1
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			ReDim preserve FORM_QUESTION_ARRAY(question_num)		'Case "Combined AR for Certain Pops (DHS-3727)"
			Set FORM_QUESTION_ARRAY(question_num) = new form_questions
			FORM_QUESTION_ARRAY(question_num).number 				= 5
			FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "For SNAP, did you or your spouse win a cash prize from lottery or gambling of $4,250 or more, in a single game or play?"
			FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q5. Did you win a cash prize, lottery or gambling?"
			FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 5. For SNAP, did you or your spouse win a cash prize from lottery or gambling of $4,250 or more, in a single game or play?"
			FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
			FORM_QUESTION_ARRAY(question_num).answer_is_array 		= false
			FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
			FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
			FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num
			FORM_QUESTION_ARRAY(question_num).verif_verbiage 		= "Combined AR Q5 Information (Winnings)"
			FORM_QUESTION_ARRAY(question_num).detail_array_exists	= True
			FORM_QUESTION_ARRAY(question_num).detail_source			= "winnings"
			FORM_QUESTION_ARRAY(question_num).detail_button_label 	= "ADD WINNING"
			FORM_QUESTION_ARRAY(question_num).detail_interview_notes= array("")
			FORM_QUESTION_ARRAY(question_num).detail_write_in_info	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_verif_status	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_verif_notes	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_edit_btn		= array("")
			FORM_QUESTION_ARRAY(question_num).detail_edit_btn		= array(2000+question_num*10)
			FORM_QUESTION_ARRAY(question_num).detail_resident_name	= array("")
			FORM_QUESTION_ARRAY(question_num).detail_amount			= array("")
			FORM_QUESTION_ARRAY(question_num).detail_date			= array("")

			FORM_QUESTION_ARRAY(question_num).mandated 				= False
			FORM_QUESTION_ARRAY(question_num).error_info			= ""
			FORM_QUESTION_ARRAY(question_num).error_verbiage		= ""
			FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
			FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
			question_num = question_num + 1

			numb_of_quest = question_num-1
			last_page_of_questions = 4
	End Select

	' numb_of_quest = UBound(FORM_QUESTION_ARRAY)
	ReDim TEMP_INFO_ARRAY(q_last_const, numb_of_quest)
End If

