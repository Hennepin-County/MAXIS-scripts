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
	public write_in_info
	public interview_notes
	public item_info_list
	public item_note_info_list
	public item_ans_list
	public item_detail_list
	public allow_prefil
	public associated_array
	public supplemental_questions
	public entirely_blank

	public verif_status
	public verif_notes

	public guide_btn
	public verif_btn
	public prefil_btn
	public add_to_array_btn
	public remove_from_array_btn

	public dialog_page_numb
	public dialog_order
	public dialog_height


	public sub display_in_dialog(y_pos, question_yn, question_notes, question_interview_notes, addtl_question, TEMP_ARRAY)
		question_answers = ""+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"Blank"
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

		If info_type = "standard" Then
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
			Text 15, y_pos, 400, 10, sub_number & "." & sub_phrase
			Text 375, y_pos, 55, 10, "CAF Answer"
			DropListBox 415, y_pos - 5, 35, 45, question_answers, addtl_question
			y_pos = y_pos + 15
			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", verif_btn
			y_pos = y_pos + 20
		ElseIf info_type = "jobs" Then
			grp_len = 35
			for each_job = 0 to UBOUND(JOBS_ARRAY, 2)
				' If JOBS_ARRAY(jobs_employer_name, each_job) <> "" AND JOBS_ARRAY(jobs_employee_name, each_job) <> "" AND JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" AND JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then
				If JOBS_ARRAY(jobs_employer_name, each_job) <> "" OR JOBS_ARRAY(jobs_employee_name, each_job) <> "" OR JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" OR JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then grp_len = grp_len + 20
			next
			GroupBox 5, y_pos, 475, grp_len, number & "." & dialog_phrasing
			PushButton 425, y_pos, 55, 10, "ADD JOB", add_to_array_btn
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_yn
			Text 95, y_pos, 25, 10, "write-in:"
			EditBox 120, y_pos - 5, 350, 15, question_notes
			' Text 360, y_pos, 110, 10, "Q9 - Verification - " & question_9_verif_yn
			' y_pos = y_pos + 20

			' PushButton 300, 100, 75, 10, "ADD VERIFICATION", add_verif_9_btn
			' y_pos = 110
			' If JOBS_ARRAY(jobs_employee_name, 0) <> "" Then
			First_job = TRUE
			for each_job = 0 to UBOUND(JOBS_ARRAY, 2)
				' If JOBS_ARRAY(jobs_employer_name, each_job) <> "" AND JOBS_ARRAY(jobs_employee_name, each_job) <> "" AND JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" AND JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then
				If JOBS_ARRAY(jobs_employer_name, each_job) <> "" OR JOBS_ARRAY(jobs_employee_name, each_job) <> "" OR JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" OR JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then
					If First_job = TRUE Then y_pos = y_pos + 20
					First_job = FALSE
					If JOBS_ARRAY(verif_yn, each_job) = "" Then Text 15, y_pos, 395, 10, "Employer: " & JOBS_ARRAY(jobs_employer_name, each_job) & "  - Employee: " & JOBS_ARRAY(jobs_employee_name, each_job) & "   - Gross Monthly Earnings: $ " & JOBS_ARRAY(jobs_gross_monthly_earnings, each_job)
					If JOBS_ARRAY(verif_yn, each_job) <> "" Then Text 15, y_pos, 395, 10, "Employer: " & JOBS_ARRAY(jobs_employer_name, each_job) & "  - Employee: " & JOBS_ARRAY(jobs_employee_name, each_job) & "   - Gross Monthly Earnings: $ " & JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) & "   - Verification - " & JOBS_ARRAY(verif_yn, each_job)
					PushButton 450, y_pos, 20, 10, "EDIT", JOBS_ARRAY(jobs_edit_btn, each_job)
					y_pos = y_pos + 10
				End If
			next
			If First_job = TRUE Then y_pos = y_pos + 10
			y_pos = y_pos + 15
		ElseIf info_type = "single-detail" Then
			GroupBox 5, y_pos, 475, dialog_height-5, number & "." & dialog_phrasing
			y_pos = y_pos + 20
			Text 15, y_pos, 40, 10, "CAF Answer"
			DropListBox 55, y_pos - 5, 35, 45, question_answers, question_yn
			Text 95, y_pos, 50, 10, sub_phrase & ":"
			EditBox 145, y_pos - 5, 35, 15, addtl_question
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



		ElseIf info_type = "housing" or info_type = "utilities" or info_type = "assets" or info_type = "msa" Then


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
			If info_type = "msa" Then
				col_1_1 = 25
				col_1_2 = 90
				col_2_1 = 230
				col_2_2 = 295
				drplst_len = 60
				txt_len = 140
			End If

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
			y_pos = y_pos + 5

			Text 15, y_pos, 25, 10, "Write-in:"
			If verif_status = "" Then
				EditBox 40, y_pos - 5, 435, 15, question_notes
			Else
				EditBox 40, y_pos - 5, 315, 15, question_notes
				Text 360, y_pos, 110, 10, "Q14 - Verification - " & verif_status
			End If
			y_pos = y_pos + 20

			Text 15, y_pos, 60, 10, "Interview Notes:"
			EditBox 75, y_pos - 5, 320, 15, question_interview_notes
			PushButton 400, y_pos, 75, 10, "ADD VERIFICATION", verif_btn
			y_pos = y_pos + 25

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
		If info_type = "standard" or info_type = "two-part" or info_type = "single-detail" or IsArray(associated_array) = true Then
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

			If IsArray(associated_array) = true Then
				for each_job = 0 to UBOUND(associated_array, 2)
					If associated_array(jobs_employer_name, each_job) <> "" OR associated_array(jobs_employee_name, each_job) <> "" OR associated_array(jobs_gross_monthly_earnings, each_job) <> "" OR associated_array(jobs_hourly_wage, each_job) <> "" Then
						CALL write_variable_in_CASE_NOTE("    Employer: " & associated_array(jobs_employer_name, each_job) & " for " & associated_array(jobs_employee_name, each_job) & " monthly earnings $" & associated_array(jobs_gross_monthly_earnings, each_job))
						If associated_array(verif_yn, each_job) <> "" Then
							If trim(associated_array(verif_details, each_job)) = "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & associated_array(verif_yn, each_job))
							If trim(associated_array(verif_details, each_job)) <> "" Then CALL write_variable_in_CASE_NOTE("    Verification: " & associated_array(verif_yn, each_job) & ": " & associated_array(verif_details, each_job))
						End If
						If trim(associated_array(jobs_notes, each_job)) <> "" Then CALL write_variable_in_CASE_NOTE("    WriteIn Answer: " & associated_array(jobs_notes, each_job))
						If trim(associated_array(jobs_intv_notes, each_job)) <> "" Then CALL write_variable_in_CASE_NOTE("    INTVW NOTES: " & associated_array(jobs_intv_notes, each_job))
					End If
				next
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
		ElseIf info_type = "housing" or info_type = "utilities" or info_type = "assets" or info_type = "msa" Then
			If entirely_blank = false Then
				Call write_variable_in_CASE_NOTE(note_phrasing)
				CALL write_variable_in_CASE_NOTE("    CAF Answer:")

				for i = 0 to UBound(item_ans_list)
					item_ans_list(i) = left(item_ans_list(i) & "   ", 5)
				next
				If info_type = "housing" Then
					spaces_1 = "       "
					spaces_2 = "                        "
					CALL write_variable_in_CASE_NOTE(spaces_1 & "Rent - " & item_ans_list(0) &  " Rental Subsidy - " & item_ans_list(1) & "  Mortgage - " & item_ans_list(2) & "    Taxes - " & item_ans_list(3))
					CALL write_variable_in_CASE_NOTE(spaces_2 & "Assoc Fees - " & item_ans_list(4) & "Room/Board - " & item_ans_list(5)    & "Insurance - " & item_ans_list(6))
				End If
				If info_type = "utilities" Then
					CALL write_variable_in_CASE_NOTE("        Heat/AC - " & item_ans_list(0) & " Electric - " & item_ans_list(1) & " Cooking Fuel - " & item_ans_list(2))
					CALL write_variable_in_CASE_NOTE("    Water/Sewer - " & item_ans_list(3) & "  Garbage - " & item_ans_list(4) & "        Phone - " & item_ans_list(5))
					CALL write_variable_in_CASE_NOTE("    LIHEAP/Energy Assistance in past 12 months - " & item_ans_list(6))
				End If
				If info_type = "assets" Then
					CALL write_variable_in_CASE_NOTE("      Cash - " & item_ans_list(0) & " Bank Accounts - " & item_ans_list(1))
					CALL write_variable_in_CASE_NOTE("    Stocks - " & item_ans_list(2) & "      Vehicles - " & item_ans_list(3))
				End If
				If info_type = "msa" Then
					CALL write_variable_in_CASE_NOTE("    REP Payee Fees - " & item_ans_list(0) & "         Guard Fees - " & item_ans_list(1))
					CALL write_variable_in_CASE_NOTE("      Special Diet - " & item_ans_list(2) & " High Housing Costs - " & item_ans_list(3))
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

		If info_type = "standard" or info_type = "two-part" or info_type = "single-detail" or IsArray(associated_array) = true Then
			objSelection.TypeText chr(9) & "CAF Answer: " & caf_answer & vbCr
			If info_type = "two-part" Then
				objSelection.TypeText "Q "& number&"."&sub_number&". "&sub_phrase & vbCr
				objSelection.TypeText chr(9) & sub_answer & vbCr
			End If
			If info_type = "single-detail" Then
				If sub_answer <> "" Then objSelection.TypeText chr(9) & sub_note_phrase & ": " & sub_answer & vbCr
				If sub_answer = "" Then objSelection.TypeText chr(9) & sub_note_phrase & ": NONE LISTED" & vbCr
			End If
			If write_in_info <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & write_in_info & vbCr
			If verif_status <> "Mot Needed" AND verif_status <> "" Then objSelection.TypeText chr(9) & "Verification: " & verif_status & vbCr
			If verif_notes <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & verif_notes & vbCr
			If caf_answer <> "" OR trim(write_in_info) <> "" Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
			If interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & interview_notes & vbCR
		ElseIf info_type = "unea" or info_type = "housing"  Then

			all_the_tables = UBound(TABLE_ARRAY) + 1
			ReDim Preserve TABLE_ARRAY(all_the_tables)
			Set objRange = objSelection.Range					'range is needed to create tables
			objDoc.Tables.Add objRange, 5, 1					'This sets the rows and columns needed row then column'
			set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
			table_count = table_count + 1

			'note that this table does not use an autoformat - which is why there are no borders on this table.'

			If info_type = "unea" Then
				TABLE_ARRAY(array_counters).Columns(1).Width = 500
				numb_of_rows = 5
				number_of_columns = 6
				For row = 1 to 4
					TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 6, TRUE

					TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 75, 2
					TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 100, 2
					TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 75, 2
					TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 75, 2
					TABLE_ARRAY(array_counters).Cell(row, 5).SetWidth 100, 2
					TABLE_ARRAY(array_counters).Cell(row, 6).SetWidth 75, 2
				Next
				TABLE_ARRAY(array_counters).Rows(5).Cells.Split 1, 3, TRUE

				TABLE_ARRAY(array_counters).Cell(5, 1).SetWidth 75, 2
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



			ElseIf info_type = "assets" Then



			ElseIf info_type = "msa" Then

			End If

			row = 1
			col = 1
			for i = 0 to UBound(item_ans_list)
				TABLE_ARRAY(array_counters).Cell(row, col).Range.Text = item_ans_list(i)
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
			next
			' TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = item_ans_list(0)
			' TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = item_note_info_list(0)
			' TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = "$ " & item_detail_list(0)
			' TABLE_ARRAY(array_counters).Cell(1, 4).Range.Text = item_ans_list(1)
			' TABLE_ARRAY(array_counters).Cell(1, 5).Range.Text = item_note_info_list(1)
			' TABLE_ARRAY(array_counters).Cell(1, 6).Range.Text = "$ " & item_detail_list(1)

			' TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text = item_ans_list(2)
			' TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text = item_note_info_list(2)
			' TABLE_ARRAY(array_counters).Cell(2, 3).Range.Text = "$ " & item_detail_list(2)
			' TABLE_ARRAY(array_counters).Cell(2, 4).Range.Text = item_ans_list(3)
			' TABLE_ARRAY(array_counters).Cell(2, 5).Range.Text = item_note_info_list(3)
			' TABLE_ARRAY(array_counters).Cell(2, 6).Range.Text = "$ " & item_detail_list(3)

			' TABLE_ARRAY(array_counters).Cell(3, 1).Range.Text = item_ans_list(4)
			' TABLE_ARRAY(array_counters).Cell(3, 2).Range.Text = item_note_info_list(4)
			' TABLE_ARRAY(array_counters).Cell(3, 3).Range.Text = "$ " & item_detail_list(4)
			' TABLE_ARRAY(array_counters).Cell(3, 4).Range.Text = item_ans_list(5)
			' TABLE_ARRAY(array_counters).Cell(3, 5).Range.Text = item_note_info_list(5)
			' TABLE_ARRAY(array_counters).Cell(3, 6).Range.Text = "$ " & item_detail_list(5)

			' TABLE_ARRAY(array_counters).Cell(4, 1).Range.Text = item_ans_list(6)
			' TABLE_ARRAY(array_counters).Cell(4, 2).Range.Text = item_note_info_list(6)
			' TABLE_ARRAY(array_counters).Cell(4, 3).Range.Text = "$ " & item_detail_list(6)
			' TABLE_ARRAY(array_counters).Cell(4, 4).Range.Text = item_ans_list(7)
			' TABLE_ARRAY(array_counters).Cell(4, 5).Range.Text = item_note_info_list(7)
			' TABLE_ARRAY(array_counters).Cell(4, 6).Range.Text = "$ " & item_detail_list(7)

			' TABLE_ARRAY(array_counters).Cell(5, 1).Range.Text = item_ans_list(8)
			' TABLE_ARRAY(array_counters).Cell(5, 2).Range.Text = item_note_info_list(8)
			' TABLE_ARRAY(array_counters).Cell(5, 3).Range.Text = "$ " & item_detail_list(8)

			objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
			array_counters = array_counters + 1

			If write_in_info <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & write_in_info & vbCr
			If verif_status <> "Mot Needed" AND verif_status <> "" Then objSelection.TypeText chr(9) & "Verification: " & verif_status & vbCr
			If verif_notes <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & verif_notes & vbCr

			If entirely_blank = false Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
			If interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & interview_notes & vbCR
		' ElseIf info_type = "housing" Then

		' 	all_the_tables = UBound(TABLE_ARRAY) + 1
		' 	ReDim Preserve TABLE_ARRAY(all_the_tables)
		' 	Set objRange = objSelection.Range					'range is needed to create tables
		' 	objDoc.Tables.Add objRange, 4, 1					'This sets the rows and columns needed row then column'
		' 	set TABLE_ARRAY(array_counters) = objDoc.Tables(table_count)		'Creates the table with the specific index'
		' 	table_count = table_count + 1

		' 	'note that this table does not use an autoformat - which is why there are no borders on this table.'
		' 	TABLE_ARRAY(array_counters).Columns(1).Width = 520

		' 	For row = 1 to 3
		' 		TABLE_ARRAY(array_counters).Rows(row).Cells.Split 1, 4, TRUE

		' 		TABLE_ARRAY(array_counters).Cell(row, 1).SetWidth 90, 2
		' 		TABLE_ARRAY(array_counters).Cell(row, 2).SetWidth 170, 2
		' 		TABLE_ARRAY(array_counters).Cell(row, 3).SetWidth 90, 2
		' 		TABLE_ARRAY(array_counters).Cell(row, 4).SetWidth 170, 2
		' 	Next
		' 	TABLE_ARRAY(array_counters).Rows(4).Cells.Split 1, 2, TRUE

		' 	TABLE_ARRAY(array_counters).Cell(4, 1).SetWidth 90, 2
		' 	TABLE_ARRAY(array_counters).Cell(4, 2).SetWidth 430, 2

		' 	TABLE_ARRAY(array_counters).Cell(1, 1).Range.Text = item_ans_list(0)
		' 	TABLE_ARRAY(array_counters).Cell(1, 2).Range.Text = "Rent (include mobile home lot rental)"
		' 	TABLE_ARRAY(array_counters).Cell(1, 3).Range.Text = item_ans_list(1)
		' 	TABLE_ARRAY(array_counters).Cell(1, 4).Range.Text = "Rent or Section 8 subsidy"

		' 	TABLE_ARRAY(array_counters).Cell(2, 1).Range.Text = item_ans_list(2)
		' 	TABLE_ARRAY(array_counters).Cell(2, 2).Range.Text = "Mortgage/contract for deed payment"
		' 	TABLE_ARRAY(array_counters).Cell(2, 3).Range.Text = item_ans_list(3)
		' 	TABLE_ARRAY(array_counters).Cell(2, 4).Range.Text = "Association fees"

		' 	TABLE_ARRAY(array_counters).Cell(3, 1).Range.Text = item_ans_list(4)
		' 	TABLE_ARRAY(array_counters).Cell(3, 2).Range.Text = "Homeowner's insurance (if not included in mortgage) "
		' 	TABLE_ARRAY(array_counters).Cell(3, 3).Range.Text = item_ans_list(5)
		' 	TABLE_ARRAY(array_counters).Cell(3, 4).Range.Text = "Room and/or board"

		' 	TABLE_ARRAY(array_counters).Cell(4, 1).Range.Text = item_ans_list(6)
		' 	TABLE_ARRAY(array_counters).Cell(4, 2).Range.Text = "Real estate taxes (if not included in mortgage)"

		' 	objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
		' 	array_counters = array_counters + 1

		' 	If write_in_info <> "" Then objSelection.TypeText chr(9) & "CAF Info Write-In: " & write_in_info & vbCr
		' 	If verif_status <> "Mot Needed" AND verif_status <> "" Then objSelection.TypeText chr(9) & "Verification: " & verif_status & vbCr
		' 	If verif_notes <> "" Then objSelection.TypeText chr(9) & chr(9) & "Details: " & verif_notes & vbCr

		' 	If entirely_blank = false Then objSelection.TypeText chr(9) & "CAF Confirmed during the Interview" & vbCR
		' 	If interview_notes <> "" Then objSelection.TypeText chr(9) & "Notes from Interview: " & interview_notes & vbCR
		ElseIf info_type = "utilities" Then



		ElseIf info_type = "assets" Then



		ElseIf info_type = "msa" Then



		End If
	end sub

end class
		' ReDim preserve FORM_QUESTION_ARRAY(question_num)
		' Set FORM_QUESTION_ARRAY(question_num) = new form_questions
		' FORM_QUESTION_ARRAY(question_num).number 				= 1
		' FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= ""
		' FORM_QUESTION_ARRAY(question_num).note_phrasing			= ""
		' FORM_QUESTION_ARRAY(question_num).doc_phrasing			= ""
		' FORM_QUESTION_ARRAY(question_num).info_type				= ""
		' FORM_QUESTION_ARRAY(question_num).caf_answer 			= ""
		' FORM_QUESTION_ARRAY(question_num).write_in_info 		= ""
		' FORM_QUESTION_ARRAY(question_num).interview_notes		= ""

		' FORM_QUESTION_ARRAY(question_num).verif_status 			= ""
		' FORM_QUESTION_ARRAY(question_num).verif_notes			= ""

		' FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
		' FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num

		' FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 1
		' FORM_QUESTION_ARRAY(question_num).dialog_order 			= 1
		' FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
		' question_num = question_num + 1



CAF_form = "CAF (DHS-5223)"

Const end_of_doc = 6			'This is for word document ennumeration

question_num = 0
Dim FORM_QUESTION_ARRAY()
ReDim FORM_QUESTION_ARRAY(0)

numb_of_quest = 0
last_page_of_questions = 4
Select Case CAF_form
	Case "CAF (DHS-5223)"

		ReDim preserve FORM_QUESTION_ARRAY(question_num)
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

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
		FORM_QUESTION_ARRAY(question_num).dialog_order 			= 1
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
		question_num = question_num + 1

		ReDim preserve FORM_QUESTION_ARRAY(question_num)
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

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
		FORM_QUESTION_ARRAY(question_num).dialog_order 			= 2
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
		question_num = question_num + 1

		ReDim preserve FORM_QUESTION_ARRAY(question_num)
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

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
		FORM_QUESTION_ARRAY(question_num).dialog_order 			= 3
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
		question_num = question_num + 1

		ReDim preserve FORM_QUESTION_ARRAY(question_num)
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

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
		FORM_QUESTION_ARRAY(question_num).dialog_order 			= 4
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
		question_num = question_num + 1

		ReDim preserve FORM_QUESTION_ARRAY(question_num)
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

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
		FORM_QUESTION_ARRAY(question_num).dialog_order 			= 5
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
		question_num = question_num + 1

		ReDim preserve FORM_QUESTION_ARRAY(question_num)
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

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
		FORM_QUESTION_ARRAY(question_num).dialog_order 			= 6
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
		question_num = question_num + 1

		ReDim preserve FORM_QUESTION_ARRAY(question_num)
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

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 5
		FORM_QUESTION_ARRAY(question_num).dialog_order 			= 1
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
		question_num = question_num + 1

		ReDim preserve FORM_QUESTION_ARRAY(question_num)			'HAS SUB QUESTION
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

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 5
		FORM_QUESTION_ARRAY(question_num).dialog_order 			= 2
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 75
		question_num = question_num + 1

		ReDim preserve FORM_QUESTION_ARRAY(question_num)
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
		FORM_QUESTION_ARRAY(question_num).add_to_array_btn		= 3000+question_num
		FORM_QUESTION_ARRAY(question_num).remove_from_array_btn	= 3500+question_num
		FORM_QUESTION_ARRAY(question_num).associated_array 		= JOBS_ARRAY

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 5
		FORM_QUESTION_ARRAY(question_num).dialog_order 			= 3
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 40
		question_num = question_num + 1

		ReDim preserve FORM_QUESTION_ARRAY(question_num)
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

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 5
		FORM_QUESTION_ARRAY(question_num).dialog_order 			= 4
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
		question_num = question_num + 1

		ReDim preserve FORM_QUESTION_ARRAY(question_num)
		Set FORM_QUESTION_ARRAY(question_num) = new form_questions
		FORM_QUESTION_ARRAY(question_num).number 				= 11
		FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Do you expect any changes in income, expenses or work hours?"
		FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q11.Do you expect any changes in income, expenses, or work hours?"
		FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 11. Do you expect any changes in income, expenses or work hours?"
		FORM_QUESTION_ARRAY(question_num).info_type				= "standard"
		FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
		FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
		FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 5
		FORM_QUESTION_ARRAY(question_num).dialog_order 			= 5
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
		question_num = question_num + 1

		ReDim preserve FORM_QUESTION_ARRAY(question_num)
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
		FORM_QUESTION_ARRAY(question_num).prefil_btn			= 2000+question_num

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 6
		FORM_QUESTION_ARRAY(question_num).dialog_order 			= 1
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 130
		question_num = question_num + 1

		ReDim preserve FORM_QUESTION_ARRAY(question_num)
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

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 6
		FORM_QUESTION_ARRAY(question_num).dialog_order 			= 2
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
		question_num = question_num + 1

		ReDim preserve FORM_QUESTION_ARRAY(question_num)
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
		FORM_QUESTION_ARRAY(question_num).prefil_btn			= 2000+question_num

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 7
		FORM_QUESTION_ARRAY(question_num).dialog_order 			= 1
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 135
		question_num = question_num + 1

		ReDim preserve FORM_QUESTION_ARRAY(question_num)
		Set FORM_QUESTION_ARRAY(question_num) = new form_questions
		FORM_QUESTION_ARRAY(question_num).number 				= 15
		FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does your household have the following utility expenses any time during the year?"
		FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q15.Are there any of the following utility expenses?"
		FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 15. Does your household have the following utility expenses any time during the year?"
		FORM_QUESTION_ARRAY(question_num).info_type				= "utilities"
		FORM_QUESTION_ARRAY(question_num).answer_is_array 		= true
		FORM_QUESTION_ARRAY(question_num).item_info_list 		= array("Heating/air conditioning", "Electricity", "Cooking fuel", "Water and sewer", "Garbage removal", "Pone/cell phone", "Did you or anyone in your houehold receive LIHEAP (energy assistance) for more than $20 in the past 12 months?)")
		FORM_QUESTION_ARRAY(question_num).item_note_info_list	= array()
		FORM_QUESTION_ARRAY(question_num).item_ans_list			= array("", "", "", "", "", "", "")
		FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
		FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
		FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 7
		FORM_QUESTION_ARRAY(question_num).dialog_order 			= 2
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 120
		question_num = question_num + 1

		ReDim preserve FORM_QUESTION_ARRAY(question_num)
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

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 8
		FORM_QUESTION_ARRAY(question_num).dialog_order 			= 1
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
		question_num = question_num + 1

		ReDim preserve FORM_QUESTION_ARRAY(question_num)
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

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 8
		FORM_QUESTION_ARRAY(question_num).dialog_order 			= 2
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
		question_num = question_num + 1

		ReDim preserve FORM_QUESTION_ARRAY(question_num)
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

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 8
		FORM_QUESTION_ARRAY(question_num).dialog_order 			= 3
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
		question_num = question_num + 1

		ReDim preserve FORM_QUESTION_ARRAY(question_num)
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

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 8
		FORM_QUESTION_ARRAY(question_num).dialog_order 			= 4
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
		question_num = question_num + 1

		ReDim preserve FORM_QUESTION_ARRAY(question_num)
		Set FORM_QUESTION_ARRAY(question_num) = new form_questions
		FORM_QUESTION_ARRAY(question_num).number 				= 20
		FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "Does anyone in the household own, or is anyone buying, any of the following?"
		FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q20.Does anyone own or is anyone buying any of the following:"
		FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 20. Does anyone in the household own, or is anyone buying, any of the following? Check yes or no for each item. "
		FORM_QUESTION_ARRAY(question_num).info_type				= "assets"
		FORM_QUESTION_ARRAY(question_num).answer_is_array 		= true
		FORM_QUESTION_ARRAY(question_num).item_info_list 		= array("Cash", "Bank accounts (savings, checking, debit card, etc)", "Stocks, bonds, annuities, 401k, etc", "Vehicles (cars, trucks, motorcycles, campers, trailers)")
		FORM_QUESTION_ARRAY(question_num).item_note_info_list	= array()
		FORM_QUESTION_ARRAY(question_num).item_ans_list			= array("", "", "", "")
		FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
		FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
		FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 8
		FORM_QUESTION_ARRAY(question_num).dialog_order 			= 5
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 105
		question_num = question_num + 1

		ReDim preserve FORM_QUESTION_ARRAY(question_num)
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

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 9
		FORM_QUESTION_ARRAY(question_num).dialog_order 			= 1
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
		question_num = question_num + 1

		ReDim preserve FORM_QUESTION_ARRAY(question_num)
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

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 9
		FORM_QUESTION_ARRAY(question_num).dialog_order 			= 2
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
		question_num = question_num + 1

		ReDim preserve FORM_QUESTION_ARRAY(question_num)
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

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 9
		FORM_QUESTION_ARRAY(question_num).dialog_order 			= 3
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
		question_num = question_num + 1

		ReDim preserve FORM_QUESTION_ARRAY(question_num)
		Set FORM_QUESTION_ARRAY(question_num) = new form_questions
		FORM_QUESTION_ARRAY(question_num).number 				= 24
		FORM_QUESTION_ARRAY(question_num).dialog_phrasing		= "For MSA recipients only: Does anyone in the household have any of the following expenses?"
		FORM_QUESTION_ARRAY(question_num).note_phrasing			= "Q24.Does anyone have any of the following expenses? (MSA ONLY)"
		FORM_QUESTION_ARRAY(question_num).doc_phrasing			= "Q 24. For MSA recipients only: Does anyone in the household have any of the following expenses?"
		FORM_QUESTION_ARRAY(question_num).info_type				= "msa"
		FORM_QUESTION_ARRAY(question_num).answer_is_array 		= true
		FORM_QUESTION_ARRAY(question_num).item_info_list 		= array("Representative Payee fees", "Guardian Conservator fees", "Physician-perscribed special diet", "High housing costs")
		FORM_QUESTION_ARRAY(question_num).item_note_info_list	= array()
		FORM_QUESTION_ARRAY(question_num).item_ans_list			= array("", "", "", "")
		FORM_QUESTION_ARRAY(question_num).supplemental_questions= array("")
		FORM_QUESTION_ARRAY(question_num).guide_btn 			= 500+question_num
		FORM_QUESTION_ARRAY(question_num).verif_btn 			= 1000+question_num

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 9
		FORM_QUESTION_ARRAY(question_num).dialog_order 			= 4
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 105
		question_num = question_num + 1

		numb_of_quest = question_num-1
		last_page_of_questions = 9


	Case "HUF (DHS-8107)"
	Case "SNAP App for Srs (DHS-5223F)"
	Case "MNbenefits"
	Case "Combined AR for Certain Pops (DHS-3727)"
End Select

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

Dim TEMP_HOUSING_ARRAY(6)
Dim TEMP_UTILITIES_ARRAY(6)
DIM TEMP_ASSETS_ARRAY(3)
DIM TEMP_MSA_ARRAY(3)

const jobs_employee_name 			= 0
const jobs_hourly_wage 				= 1
const jobs_gross_monthly_earnings	= 2
const jobs_employer_name 			= 3
const jobs_edit_btn					= 4
const jobs_intv_notes				= 5
const verif_yn						= 6
const verif_details					= 7
const jobs_notes 					= 8
Dim JOBS_ARRAY
ReDim JOBS_ARRAY(jobs_notes, 0)
Dim TABLE_ARRAY


const form_yn_const			= 0
const fomr_second_yn_const	= 1
const form_write_in_const	= 2
const intv_notes_const 		= 3
const verif_yn_const 		= 4
const verif_notes_const		= 5
const q_last_const			= 10

' numb_of_quest = UBound(FORM_QUESTION_ARRAY)
Dim TEMP_INFO_ARRAY()
ReDim TEMP_INFO_ARRAY(q_last_const, numb_of_quest)

MAXIS_case_number = "344839"
MsgBox "CAREFUL! This will CASE/NOTE in " & MAXIS_case_number & " without any real warning." & vbCr & vbCr & "USE IN TRAINING REGION."

page_display = 1
Do
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 555, 385, "Full Interview Questions"

		ButtonGroup ButtonPressed
			If page_display = 1 Then
				ComboBox 120, 10, 205, 45, all_the_clients+chr(9)+who_are_we_completing_the_interview_with, who_are_we_completing_the_interview_with
				ComboBox 120, 30, 75, 45, "Select or Type"+chr(9)+"Phone"+chr(9)+"In Office"+chr(9)+how_are_we_completing_the_interview, how_are_we_completing_the_interview
				EditBox 120, 50, 50, 15, interview_date
				ComboBox 120, 70, 340, 45, "No Interpreter Used"+chr(9)+"Language Line Interpreter Used"+chr(9)+"Interpreter through Henn Co. OMS (Office of Multi-Cultural Services)"+chr(9)+"Interviewer speaks Resident Language"+chr(9)+interpreter_information, interpreter_information
				ComboBox 120, 90, 205, 45, "English"+chr(9)+"Somali"+chr(9)+"Spanish"+chr(9)+"Hmong"+chr(9)+"Russian"+chr(9)+"Oromo"+chr(9)+"Vietnamese"+chr(9)+interpreter_language, interpreter_language
				PushButton 330, 90, 120, 15, "Open Interpreter Services Link", interpreter_servicves_btn
				EditBox 120, 110, 340, 15, arep_interview_id_information
				EditBox 10, 155, 450, 15, non_applicant_interview_info
			ElseIf page_display = 2 Then
				GroupBox 10, 35, 375, 95, "Residence Address"
				Text 20, 55, 45, 10, "House/Street"
				Text 45, 75, 20, 10, "City"
				Text 185, 75, 20, 10, "State"
				Text 325, 75, 15, 10, "Zip"
				Text 20, 95, 100, 10, "Do you live on a Reservation?"
				Text 180, 95, 60, 10, "If yes, which one?"
				Text 20, 115, 100, 10, "Resident Indicates Homeless:"
				Text 185, 115, 60, 10, "Living Situation?"
				GroupBox 10, 135, 375, 70, "Mailing Address"
				Text 20, 165, 45, 10, "House/Street"
				Text 45, 185, 20, 10, "City"
				Text 185, 185, 20, 10, "State"
				Text 325, 185, 15, 10, "Zip"
				GroupBox 10, 210, 235, 90, "Phone Number"
				Text 20, 225, 50, 10, "Number"
				Text 125, 225, 25, 10, "Type"
				Text 255, 225, 60, 10, "Date of Change:"
				Text 255, 245, 75, 10, "County of Residence:"
			ElseIf page_display = 3 Then
				Text 70, 35, 50, 10, "Last Name"
				Text 165, 35, 50, 10, "First Name"
				Text 245, 35, 50, 10, "Middle Name"
				Text 300, 35, 50, 10, "Other Names"
				Text 70, 65, 55, 10, "Soc Sec Number"
				Text 145, 65, 45, 10, "Date of Birth"
				Text 220, 65, 45, 10, "Gender"
				Text 275, 65, 90, 10, "Relationship to MEMB 01"
				Text 370, 65, 50, 10, "Marital Status"
				Text 70, 95, 75, 10, "Last Grade Completed"
				Text 185, 95, 55, 10, "Moved to MN on"
				Text 260, 95, 65, 10, "Moved to MN from"
				Text 400, 95, 75, 10, "US Citizen or National"
				Text 70, 125, 40, 10, "Interpreter?"
				Text 140, 125, 95, 10, "Preferred Spoken Language"
				Text 140, 155, 95, 10, "Preferred Written Language"
				Text 70, 175, 65, 10, "Identity Verification"
				GroupBox 325, 125, 155, 100, "Demographics"
				Text 330, 135, 35, 10, "Hispanic?"
				Text 330, 160, 50, 10, "Race"
				Text 70, 200, 145, 10, "Which programs is this person requesting?"
				Text 70, 255, 80, 10, "Intends to reside in MN"
				Text 155, 255, 65, 10, "Immigration Status"
				Text 365, 255, 50, 10, "Sponsor?"
				Text 70, 285, 50, 10, "Verification"
				Text 155, 285, 65, 10, "Verification Details"
				Text 70, 315, 50, 10, "Notes:"
			ElseIf page_display >= 4 or page_display <= last_page_of_questions Then
				' display_count = 1
				y_pos = 10
				For quest = 0 to UBound(FORM_QUESTION_ARRAY)
					If FORM_QUESTION_ARRAY(quest).dialog_page_numb = page_display Then
						' If FORM_QUESTION_ARRAY(quest).dialog_order = display_count Then
						If FORM_QUESTION_ARRAY(quest).answer_is_array = false Then call FORM_QUESTION_ARRAY(quest).display_in_dialog(y_pos, TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(fomr_second_yn_const, quest), "")
						If FORM_QUESTION_ARRAY(quest).answer_is_array = true  Then
							If FORM_QUESTION_ARRAY(quest).info_type = "unea" Then call FORM_QUESTION_ARRAY(quest).display_in_dialog(y_pos, TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(fomr_second_yn_const, quest), "")
							If FORM_QUESTION_ARRAY(quest).info_type = "housing" Then call FORM_QUESTION_ARRAY(quest).display_in_dialog(y_pos, TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(fomr_second_yn_const, quest), TEMP_HOUSING_ARRAY)
							If FORM_QUESTION_ARRAY(quest).info_type = "utilities" Then call FORM_QUESTION_ARRAY(quest).display_in_dialog(y_pos, TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(fomr_second_yn_const, quest), TEMP_UTILITIES_ARRAY)
							If FORM_QUESTION_ARRAY(quest).info_type = "assets" Then call FORM_QUESTION_ARRAY(quest).display_in_dialog(y_pos, TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(fomr_second_yn_const, quest), TEMP_ASSETS_ARRAY)
							If FORM_QUESTION_ARRAY(quest).info_type = "msa" Then call FORM_QUESTION_ARRAY(quest).display_in_dialog(y_pos, TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(fomr_second_yn_const, quest), TEMP_MSA_ARRAY)
						End If
						' y_pos = y_pos + FORM_QUESTION_ARRAY(quest).dialog_height
						' MsgBox "y_pos - " & y_pos
						' 	display_count = display_count + 1
						' End If
					End If
				Next
			' ElseIf page_display = 5 Then
			' 	GroupBox 10, 35, 375, 95, "PAGE 5"

			End If

			Text 485, 5, 75, 10, "---   DIALOGS   ---"
			Text 485, 17, 10, 10, "1"
			Text 485, 32, 10, 10, "2"
			Text 485, 47, 10, 10, "3"
			Text 485, 62, 10, 10, "4"
			Text 485, 77, 10, 10, "5"
			Text 485, 92, 10, 10, "6"
			Text 485, 107, 10, 10, "7"
			Text 485, 122, 10, 10, "8"
			Text 485, 137, 10, 10, "9"

			If page_display <> 1 Then PushButton 495, 15, 55, 13, "INTVW / CAF 1", caf_page_one_btn
			If page_display <> 2 Then PushButton 495, 30, 55, 13, "CAF ADDR", caf_addr_btn
			If page_display <> 3 Then PushButton 495, 45, 55, 13, "CAF MEMBs", caf_membs_btn
			If page_display <> 4 Then PushButton 495, 60, 55, 13, "Q. 1 - 6", caf_q_1_6_btn
			If page_display <> 5 Then PushButton 495, 75, 55, 13, "Q. 7 - 11", caf_q_7_11_btn
			If page_display <> 6 Then PushButton 495, 90, 55, 13, "Q. 12 - 13", caf_q_12_13_btn
			If page_display <> 7 Then PushButton 495, 105, 55, 13, "Q. 14 - 15", caf_q_14_15_btn
			If page_display <> 8 Then PushButton 495, 120, 55, 13, "Q. 16 - 20", caf_q_16_20_btn
			If page_display <> 9 Then PushButton 495, 135, 55, 13, "Q. 21 - 24", caf_q_21_24_btn


	EndDialog


	err_msg = "LOOP"

	dialog Dialog1
	cancel_without_confirmation
	For quest = 0 to UBound(FORM_QUESTION_ARRAY)
		If FORM_QUESTION_ARRAY(quest).answer_is_array = false Then call FORM_QUESTION_ARRAY(quest).store_dialog_entry(TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(fomr_second_yn_const, quest), "")
		If FORM_QUESTION_ARRAY(quest).answer_is_array = true Then
			If FORM_QUESTION_ARRAY(quest).info_type = "unea" Then call FORM_QUESTION_ARRAY(quest).store_dialog_entry(TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(fomr_second_yn_const, quest), "")
			If FORM_QUESTION_ARRAY(quest).info_type = "housing" Then call FORM_QUESTION_ARRAY(quest).store_dialog_entry(TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(fomr_second_yn_const, quest), TEMP_HOUSING_ARRAY)
			If FORM_QUESTION_ARRAY(quest).info_type = "utilities" Then call FORM_QUESTION_ARRAY(quest).store_dialog_entry(TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(fomr_second_yn_const, quest), TEMP_UTILITIES_ARRAY)
			If FORM_QUESTION_ARRAY(quest).info_type = "assets" Then call FORM_QUESTION_ARRAY(quest).store_dialog_entry(TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(fomr_second_yn_const, quest), TEMP_ASSETS_ARRAY)
			If FORM_QUESTION_ARRAY(quest).info_type = "msa" Then call FORM_QUESTION_ARRAY(quest).store_dialog_entry(TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(fomr_second_yn_const, quest), TEMP_MSA_ARRAY)
		End If

		' If FORM_QUESTION_ARRAY(quest).info_type = "housing" Then
		' 	MsgBox "Rent answer: " & FORM_QUESTION_ARRAY(quest).item_ans_list(0)
		' End If
	Next

	For quest = 0 to UBound(FORM_QUESTION_ARRAY)
		If ButtonPressed = FORM_QUESTION_ARRAY(quest).verif_btn Then
			call FORM_QUESTION_ARRAY(quest).capture_verif_detail()
		End If
	Next

	If ButtonPressed = caf_page_one_btn Then page_display = 1
	If ButtonPressed = caf_addr_btn Then page_display = 2
	If ButtonPressed = caf_membs_btn Then page_display = 3
	If ButtonPressed = caf_q_1_6_btn Then page_display = 4
	If ButtonPressed = caf_q_7_11_btn Then page_display = 5
	If ButtonPressed = caf_q_12_13_btn Then page_display = 6
	If ButtonPressed = caf_q_14_15_btn Then page_display = 7
	If ButtonPressed = caf_q_16_20_btn Then page_display = 8
	If ButtonPressed = caf_q_21_24_btn Then page_display = 9

	If ButtonPressed = -1 Then err_msg = ""

Loop until err_msg = ""


'****writing the word document
Set objWord = CreateObject("Word.Application")

'Adding all of the information in the dialogs into a Word Document
If no_case_number_checkbox = checked Then objWord.Caption = "CAF Form Details - NEW CASE"
If no_case_number_checkbox = unchecked Then objWord.Caption = "CAF Form Details - CASE #" & MAXIS_case_number			'Title of the document
' objWord.Visible = True														'Let the worker see the document
objWord.Visible = True 														'The worker should NOT see the docuement
'allow certain workers to see the document
' If user_ID_for_validation = "WFA168" or user_ID_for_validation = "LILE002" Then objWord.Visible = True

Set objDoc = objWord.Documents.Add()										'Start a new document
Set objSelection = objWord.Selection										'This is kind of the 'inside' of the document

objSelection.Font.Name = "Arial"											'Setting the font before typing
objSelection.Font.Size = "16"
objSelection.Font.Bold = TRUE
objSelection.TypeText "NOTES on INTERVIEW"
objSelection.TypeParagraph()
objSelection.Font.Size = "14"
objSelection.Font.Bold = FALSE

If MAXIS_case_number <> "" Then objSelection.TypeText "Case Number: " & MAXIS_case_number & vbCR			'General case information


Set objRange = objSelection.Range					'range is needed to create tables
objDoc.Tables.Add objRange, 7, 1					'This sets the rows and columns needed row then column
'This table starts with 1 column - other columns are added after we split some of the cells
set objProgStatusTable = objDoc.Tables(1)		'Creates the table with the specific index'

objProgStatusTable.AutoFormat(16)							'This adds the borders to the table and formats it

objProgStatusTable.Cell(1, 1).SetHeight 15, 2
for row = 2 to 7
	objProgStatusTable.Cell(row, 1).SetHeight 12, 2			'setting the heights of the rows
Next

for row = 1 to 7
	objProgStatusTable.Rows(row).Cells.Split 1, 2, TRUE
Next

objProgStatusTable.Columns(1).Width = 150					'This sets the width of the table.
objProgStatusTable.Columns(2).Width = 200					'This sets the width of the table.
' objProgStatusTable.Columns(3).Width = 150					'This sets the width of the table.

'Now going to each cell and setting teh font size
objProgStatusTable.Cell(1, 1).Range.Font.Size = 11
objProgStatusTable.Cell(1, 2).Range.Font.Size = 11
For row = 2 to 7
	objProgStatusTable.Cell(row, 1).Range.Font.Size = 9
	objProgStatusTable.Cell(row, 2).Range.Font.Size = 9
Next

' objProgStatusTable.Cell(row, col).Range.Text =

objProgStatusTable.Cell(1, 1).Range.Text = "Program"
objProgStatusTable.Cell(1, 2).Range.Text = "Status"
' objProgStatusTable.Cell(1, 3).Range.Text = "Detail"

objProgStatusTable.Cell(2, 1).Range.Text = "SNAP"
objProgStatusTable.Cell(2, 2).Range.Text = "ACTIVE"

objProgStatusTable.Cell(3, 1).Range.Text = "CASH 1"
objProgStatusTable.Cell(3, 2).Range.Text = "INACTIVE"

objProgStatusTable.Cell(4, 1).Range.Text = "CASH 2"
objProgStatusTable.Cell(4, 2).Range.Text = "INACTIVE"

objProgStatusTable.Cell(5, 1).Range.Text = "GRH"
objProgStatusTable.Cell(5, 2).Range.Text = "INACTIVE"

objProgStatusTable.Cell(6, 1).Range.Text = "MA"
objProgStatusTable.Cell(6, 2).Range.Text = "INACTIVE"

objProgStatusTable.Cell(7, 1).Range.Text = "MSP"
objProgStatusTable.Cell(7, 2).Range.Text = "INACTIVE"

objSelection.EndKey end_of_doc						'this sets the cursor to the end of the document for more writing
objSelection.TypeText vbCr

table_count = 1			'table index variable
ReDim TABLE_ARRAY(0)			'This creates the table array for if there is only one person listed on the CAF
array_counters = 0		'the incrementer for the table array'


For each_question = 0 to UBound(FORM_QUESTION_ARRAY)
	FORM_QUESTION_ARRAY(each_question).add_to_wif()
Next


Call start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("TRIAL INTERVIEW")

Call write_variable_in_CASE_NOTE("Interview Date")

CALL write_variable_in_CASE_NOTE("-----  CAF Information and Notes -----")

For each_question = 0 to UBound(FORM_QUESTION_ARRAY)
	FORM_QUESTION_ARRAY(each_question).enter_case_note()
Next

Call write_variable_in_CASE_NOTE("SCRIPT WRITER")
