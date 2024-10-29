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
const jobs_employee_name 			= 0
const jobs_hourly_wage 				= 1
const jobs_gross_monthly_earnings	= 2
const jobs_employer_name 			= 3
const jobs_edit_btn					= 4
const jobs_intv_notes				= 5
const verif_yn						= 6
const verif_details					= 7
const jobs_notes 					= 8
Dim JOBS_ARRAY()
ReDim JOBS_ARRAY(jobs_notes, 0)
Dim TABLE_ARRAY


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

	public guide_btn
	public verif_btn
	public prefil_btn
	public add_to_array_btn
	public edit_in_array_btn

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
			' MsgBox " is detail_interview_notes an array? - " & IsArray(detail_interview_notes)
			for each_item = 0 to UBOUND(detail_interview_notes)
				' If associated_array(jobs_employer_name, each_job) <> "" AND associated_array(jobs_employee_name, each_job) <> "" AND associated_array(jobs_gross_monthly_earnings, each_job) <> "" AND associated_array(jobs_hourly_wage, each_job) <> "" Then
				' MsgBox "employer - " & associated_array(jobs_employer_name, each_job) & vbCr & "employee - " & associated_array(jobs_employee_name, each_job) & vbCr & "earnings - " & associated_array(jobs_gross_monthly_earnings, each_job) & vbCr & "hourly wage - " & associated_array(jobs_hourly_wage, each_job)
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
				' If associated_array(jobs_employer_name, each_job) <> "" OR associated_array(jobs_employee_name, each_job) <> "" OR associated_array(jobs_gross_monthly_earnings, each_job) <> "" OR associated_array(jobs_hourly_wage, each_job) <> "" Then
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
			' Text 360, y_pos, 110, 10, "Q9 - Verification - " & question_9_verif_yn
			' y_pos = y_pos + 20

			' PushButton 300, 100, 75, 10, "ADD VERIFICATION", add_verif_9_btn
			' y_pos = 110
			' If associated_array(jobs_employee_name, 0) <> "" Then
			first_item = TRUE
			for each_item = 0 to UBOUND(detail_interview_notes)
				' If associated_array(jobs_employer_name, each_job) <> "" AND associated_array(jobs_employee_name, each_job) <> "" AND associated_array(jobs_gross_monthly_earnings, each_job) <> "" AND associated_array(jobs_hourly_wage, each_job) <> "" Then
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

			' grp_len = 35
			' for each_job = 0 to UBOUND(JOBS_ARRAY, 2)
			' 	' If JOBS_ARRAY(jobs_employer_name, each_job) <> "" AND JOBS_ARRAY(jobs_employee_name, each_job) <> "" AND JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" AND JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then
			' 	If JOBS_ARRAY(jobs_employer_name, each_job) <> "" OR JOBS_ARRAY(jobs_employee_name, each_job) <> "" OR JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" OR JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then grp_len = grp_len + 20
			' next
			' GroupBox 5, y_pos, 475, grp_len, number & "." & dialog_phrasing
			' PushButton 425, y_pos, 55, 10, "ADD JOB", add_to_array_btn
			' y_pos = y_pos + 20
			' Text 15, y_pos, 40, 10, "CAF Answer"
			' DropListBox 55, y_pos - 5, 35, 45, question_answers, question_yn
			' Text 95, y_pos, 25, 10, "write-in:"
			' EditBox 120, y_pos - 5, 350, 15, question_notes
			' ' Text 360, y_pos, 110, 10, "Q9 - Verification - " & question_9_verif_yn
			' ' y_pos = y_pos + 20

			' ' PushButton 300, 100, 75, 10, "ADD VERIFICATION", add_verif_9_btn
			' ' y_pos = 110
			' ' If JOBS_ARRAY(jobs_employee_name, 0) <> "" Then
			' First_job = TRUE
			' for each_job = 0 to UBOUND(JOBS_ARRAY, 2)
			' 	' If JOBS_ARRAY(jobs_employer_name, each_job) <> "" AND JOBS_ARRAY(jobs_employee_name, each_job) <> "" AND JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" AND JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then
			' 	If JOBS_ARRAY(jobs_employer_name, each_job) <> "" OR JOBS_ARRAY(jobs_employee_name, each_job) <> "" OR JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) <> "" OR JOBS_ARRAY(jobs_hourly_wage, each_job) <> "" Then
			' 		If First_job = TRUE Then y_pos = y_pos + 20
			' 		First_job = FALSE
			' 		If JOBS_ARRAY(verif_yn, each_job) = "" Then Text 15, y_pos, 395, 10, "Employer: " & JOBS_ARRAY(jobs_employer_name, each_job) & "  - Employee: " & JOBS_ARRAY(jobs_employee_name, each_job) & "   - Gross Monthly Earnings: $ " & JOBS_ARRAY(jobs_gross_monthly_earnings, each_job)
			' 		If JOBS_ARRAY(verif_yn, each_job) <> "" Then Text 15, y_pos, 395, 10, "Employer: " & JOBS_ARRAY(jobs_employer_name, each_job) & "  - Employee: " & JOBS_ARRAY(jobs_employee_name, each_job) & "   - Gross Monthly Earnings: $ " & JOBS_ARRAY(jobs_gross_monthly_earnings, each_job) & "   - Verification - " & JOBS_ARRAY(verif_yn, each_job)
			' 		PushButton 450, y_pos, 20, 10, "EDIT", JOBS_ARRAY(jobs_edit_btn, each_job)
			' 		y_pos = y_pos + 10
			' 	End If
			' next
			' If First_job = TRUE Then y_pos = y_pos + 10
			' y_pos = y_pos + 15
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

	' temp_hours_per_week = FORM_QUESTION_ARRAY(form_question_numb).detail_hours_per_week(selected)
	' temp_verif_notes = FORM_QUESTION_ARRAY(form_question_numb).detail_verif_notes(selected)
	' temp_value = FORM_QUESTION_ARRAY(form_question_numb).detail_value(selected)
	' temp_date = FORM_QUESTION_ARRAY(form_question_numb).detail_date(selected)
	' temp_frequency = FORM_QUESTION_ARRAY(form_question_numb).detail_frequency(selected)
	' temp_amount = FORM_QUESTION_ARRAY(form_question_numb).detail_amount(selected)
	' temp_current = FORM_QUESTION_ARRAY(form_question_numb).detail_current(selected)
	' temp_explain = FORM_QUESTION_ARRAY(form_question_numb).detail_explain(selected)

	Do

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
			Text 10, det_dlg_len-15, 110, 10, "JOB Verification - " & FORM_QUESTION_ARRAY(form_question_numb).detail_verif_status(selected)
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

form_list = "CAF (DHS-5223)"
form_list = form_list+chr(9)+"HUF (DHS-8107)"
form_list = form_list+chr(9)+"SNAP App for Srs (DHS-5223F)"
form_list = form_list+chr(9)+"MNbenefits"
form_list = form_list+chr(9)+"Combined AR for Certain Pops (DHS-3727)"

Call MAXIS_case_number_finder(MAXIS_case_number)
MAXIS_case_number = "344839"
' MsgBox "CAREFUL! This will CASE/NOTE in " & MAXIS_case_number & " without any real warning." & vbCr & vbCr & "USE IN TRAINING REGION."


Dialog1 = ""
BeginDialog Dialog1, 0, 0, 201, 80, "Dialog"
  EditBox 60, 10, 60, 15, MAXIS_case_number
  DropListBox 60, 35, 130, 45, form_list, CAF_form
  ButtonGroup ButtonPressed
    OkButton 85, 55, 50, 15
    CancelButton 140, 55, 50, 15
  Text 10, 15, 50, 10, "Case Number:"
  Text 35, 40, 20, 10, "Form:"
EndDialog

Do
	err_msg = ""

	dialog Dialog1
	cancel_without_confirmation

	call validate_MAXIS_case_number(err_msg, "*")

	If err_msg <> "" Then MsgBox "Resolve:" & vbCr & err_msg

Loop until err_msg = ""



' CAF_form = "CAF (DHS-5223)"
' CAF_form = "HUF (DHS-8107)"
' CAF_form = "SNAP App for Srs (DHS-5223F)"
' CAF_form = "MNbenefits"
' CAF_form = "Combined AR for Certain Pops (DHS-3727)"

Const end_of_doc = 6			'This is for word document ennumeration

question_num = 0
Dim FORM_QUESTION_ARRAY()
ReDim FORM_QUESTION_ARRAY(0)

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
		FORM_QUESTION_ARRAY(question_num).prefil_btn			= 2000+question_num

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
		FORM_QUESTION_ARRAY(question_num).prefil_btn			= 2000+question_num

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
		FORM_QUESTION_ARRAY(question_num).prefil_btn			= 2000+question_num

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
		FORM_QUESTION_ARRAY(question_num).prefil_btn			= 2000+question_num

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
		FORM_QUESTION_ARRAY(question_num).prefil_btn			= 2000+question_num

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
		FORM_QUESTION_ARRAY(question_num).prefil_btn			= 2000+question_num

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

		FORM_QUESTION_ARRAY(question_num).dialog_page_numb 		= 4
		FORM_QUESTION_ARRAY(question_num).dialog_height 		= 60
		question_num = question_num + 1

		numb_of_quest = question_num-1
		last_page_of_questions = 4

End Select

'TODO - DEAL with TEXT AND EMAIL question for all forms

pg_4_label = ""
pg_5_label = ""
pg_6_label = ""
pg_7_label = ""
pg_8_label = ""
pg_9_label = ""
pg_10_label = ""
pg_11_label = ""
For quest = 0 to UBound(FORM_QUESTION_ARRAY)
	If pg_4_label = "" and FORM_QUESTION_ARRAY(quest).dialog_page_numb = 4 Then pg_4_label = "Q. " & FORM_QUESTION_ARRAY(quest).number & " - "
	If pg_5_label = "" and FORM_QUESTION_ARRAY(quest).dialog_page_numb = 5 Then
		If right(pg_4_label, 1) = " " Then pg_4_label = pg_4_label & FORM_QUESTION_ARRAY(quest).number-1
		pg_5_label = "Q. " & FORM_QUESTION_ARRAY(quest).number & " - "
	End If
	If pg_6_label = "" and FORM_QUESTION_ARRAY(quest).dialog_page_numb = 6 Then
		If right(pg_5_label, 1) = " " Then pg_5_label = pg_5_label & FORM_QUESTION_ARRAY(quest).number-1
		pg_6_label = "Q. " & FORM_QUESTION_ARRAY(quest).number & " - "
	End If
	If pg_7_label = "" and FORM_QUESTION_ARRAY(quest).dialog_page_numb = 7 Then
		If right(pg_6_label, 1) = " " Then pg_6_label = pg_6_label & FORM_QUESTION_ARRAY(quest).number-1
		pg_7_label = "Q. " & FORM_QUESTION_ARRAY(quest).number & " - "
	End If
	If pg_8_label = "" and FORM_QUESTION_ARRAY(quest).dialog_page_numb = 8 Then
		If right(pg_7_label, 1) = " " Then pg_7_label = pg_7_label & FORM_QUESTION_ARRAY(quest).number-1
		pg_8_label = "Q. " & FORM_QUESTION_ARRAY(quest).number & " - "
	End If
	If pg_9_label = "" and FORM_QUESTION_ARRAY(quest).dialog_page_numb = 9 Then
		If right(pg_8_label, 1) = " " Then pg_8_label = pg_8_label & FORM_QUESTION_ARRAY(quest).number-1
		pg_9_label = "Q. " & FORM_QUESTION_ARRAY(quest).number & " - "
	End If
	If pg_10_label = "" and FORM_QUESTION_ARRAY(quest).dialog_page_numb = 10 Then
		If right(pg_9_label, 1) = " " Then pg_9_label = pg_9_label & FORM_QUESTION_ARRAY(quest).number-1
		pg_10_label = "Q. " & FORM_QUESTION_ARRAY(quest).number & " - "
	End If
	If pg_11_label = "" and FORM_QUESTION_ARRAY(quest).dialog_page_numb = 11 Then
		If right(pg_10_label, 1) = " " Then pg_10_label = pg_10_label & FORM_QUESTION_ARRAY(quest).number-1
		pg_11_label = "Q. " & FORM_QUESTION_ARRAY(quest).number & " - "
	End If

Next
If right(pg_4_label, 1) = " " Then pg_4_label = pg_4_label & FORM_QUESTION_ARRAY(numb_of_quest).number
If right(pg_5_label, 1) = " " Then pg_5_label = pg_5_label & FORM_QUESTION_ARRAY(numb_of_quest).number
If right(pg_6_label, 1) = " " Then pg_6_label = pg_6_label & FORM_QUESTION_ARRAY(numb_of_quest).number
If right(pg_7_label, 1) = " " Then pg_7_label = pg_7_label & FORM_QUESTION_ARRAY(numb_of_quest).number
If right(pg_8_label, 1) = " " Then pg_8_label = pg_8_label & FORM_QUESTION_ARRAY(numb_of_quest).number
If right(pg_9_label, 1) = " " Then pg_9_label = pg_9_label & FORM_QUESTION_ARRAY(numb_of_quest).number
If right(pg_10_label, 1) = " " Then pg_10_label = pg_10_label & FORM_QUESTION_ARRAY(numb_of_quest).number
If right(pg_11_label, 1) = " " Then pg_11_label = pg_11_label & FORM_QUESTION_ARRAY(numb_of_quest).number


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
If CAF_form = "CAF (DHS-5223)" or CAF_form = "SNAP App for Srs (DHS-5223F)" Then ReDim TEMP_HOUSING_ARRAY(5)
If CAF_form <> "CAF (DHS-5223)" and CAF_form <> "SNAP App for Srs (DHS-5223F)" Then ReDim TEMP_HOUSING_ARRAY(6)
If CAF_form = "CAF (DHS-5223)" or CAF_form = "SNAP App for Srs (DHS-5223F)" Then ReDim TEMP_UTILITIES_ARRAY(5)
If CAF_form <> "CAF (DHS-5223)" and CAF_form <> "SNAP App for Srs (DHS-5223F)" Then ReDim TEMP_UTILITIES_ARRAY(6)
DIM TEMP_ASSETS_ARRAY(4)
DIM TEMP_MSA_ARRAY(3)
DIM TEMP_STWK_ARRAY(3)



const form_yn_const			= 0
const form_second_yn_const	= 1
const form_write_in_const	= 2
const intv_notes_const 		= 3
const verif_yn_const 		= 4
const verif_notes_const		= 5
const q_last_const			= 10

' numb_of_quest = UBound(FORM_QUESTION_ARRAY)
Dim TEMP_INFO_ARRAY()
ReDim TEMP_INFO_ARRAY(q_last_const, numb_of_quest)

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
						If FORM_QUESTION_ARRAY(quest).answer_is_array = false Then call FORM_QUESTION_ARRAY(quest).display_in_dialog(y_pos, TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), "")
						If FORM_QUESTION_ARRAY(quest).answer_is_array = true  Then
							If FORM_QUESTION_ARRAY(quest).info_type = "unea" Then call FORM_QUESTION_ARRAY(quest).display_in_dialog(y_pos, TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), "")
							If FORM_QUESTION_ARRAY(quest).info_type = "housing" Then call FORM_QUESTION_ARRAY(quest).display_in_dialog(y_pos, TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), TEMP_HOUSING_ARRAY)
							If FORM_QUESTION_ARRAY(quest).info_type = "utilities" Then call FORM_QUESTION_ARRAY(quest).display_in_dialog(y_pos, TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), TEMP_UTILITIES_ARRAY)
							If FORM_QUESTION_ARRAY(quest).info_type = "assets" Then call FORM_QUESTION_ARRAY(quest).display_in_dialog(y_pos, TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), TEMP_ASSETS_ARRAY)
							If FORM_QUESTION_ARRAY(quest).info_type = "msa" Then call FORM_QUESTION_ARRAY(quest).display_in_dialog(y_pos, TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), TEMP_MSA_ARRAY)
							If FORM_QUESTION_ARRAY(quest).info_type = "stwk" Then call FORM_QUESTION_ARRAY(quest).display_in_dialog(y_pos, TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), TEMP_STWK_ARRAY)
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
			If last_page_of_questions => 5 Then Text 485, 77, 10, 10, "5"
			If last_page_of_questions => 6 Then Text 485, 92, 10, 10, "6"
			If last_page_of_questions => 7 Then Text 485, 107, 10, 10, "7"
			If last_page_of_questions => 8 Then Text 485, 122, 10, 10, "8"
			If last_page_of_questions => 9 Then Text 485, 137, 10, 10, "9"
			If last_page_of_questions => 10 Then Text 485, 152, 10, 10, "10"
			If last_page_of_questions => 11 Then Text 485, 167, 10, 10, "11"

			If page_display <> 1 Then PushButton 495, 15, 55, 13, "INTVW / CAF 1", caf_page_one_btn
			If page_display <> 2 Then PushButton 495, 30, 55, 13, "CAF ADDR", caf_addr_btn
			If page_display <> 3 Then PushButton 495, 45, 55, 13, "CAF MEMBs", caf_membs_btn
			If page_display <> 4 Then PushButton 495, 60, 55, 13, pg_4_label, caf_q_pg_4
			If page_display <> 5 and last_page_of_questions => 5 Then PushButton 495, 75, 55, 13, pg_5_label, caf_q_pg_5
			If page_display <> 6 and last_page_of_questions => 6 Then PushButton 495, 90, 55, 13, pg_6_label, caf_q_pg_6
			If page_display <> 7 and last_page_of_questions => 7 Then PushButton 495, 105, 55, 13, pg_7_label, caf_q_pg_7
			If page_display <> 8 and last_page_of_questions => 8 Then PushButton 495, 120, 55, 13, pg_8_label, caf_q_pg_8
			If page_display <> 9 and last_page_of_questions => 9 Then PushButton 495, 135, 55, 13, pg_9_label, caf_q_pg_9
			If page_display <> 10 and last_page_of_questions => 10 Then PushButton 495, 150, 55, 13, pg_10_label, caf_q_pg_10
			If page_display <> 11 and last_page_of_questions => 11 Then PushButton 495, 165, 55, 13, pg_11_label, caf_q_pg_11

			If page_display = 4 and last_page_of_questions => 4 Then Text 500, 62, 55, 10, pg_4_label
			If page_display = 5 and last_page_of_questions => 5 Then Text 500, 77, 55, 10, pg_5_label
			If page_display = 6 and last_page_of_questions => 6 Then Text 500, 92, 55, 10, pg_6_label
			If page_display = 7 and last_page_of_questions => 7 Then Text 500, 107, 55, 10, pg_7_label
			If page_display = 8 and last_page_of_questions => 8 Then Text 500, 122, 55, 10, pg_8_label
			If page_display = 9 and last_page_of_questions => 9 Then Text 500, 137, 55, 10, pg_9_label
			If page_display = 10 and last_page_of_questions => 10 Then Text 500, 152, 55, 10, pg_10_label
			If page_display = 11 and last_page_of_questions => 11 Then Text 500, 167, 55, 10, pg_11_label
	EndDialog


	err_msg = "LOOP"

	dialog Dialog1
	cancel_without_confirmation
	For quest = 0 to UBound(FORM_QUESTION_ARRAY)
		If FORM_QUESTION_ARRAY(quest).answer_is_array = false Then call FORM_QUESTION_ARRAY(quest).store_dialog_entry(TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), "")
		If FORM_QUESTION_ARRAY(quest).answer_is_array = true Then
			If FORM_QUESTION_ARRAY(quest).info_type = "unea" Then call FORM_QUESTION_ARRAY(quest).store_dialog_entry(TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), "")
			If FORM_QUESTION_ARRAY(quest).info_type = "housing" Then call FORM_QUESTION_ARRAY(quest).store_dialog_entry(TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), TEMP_HOUSING_ARRAY)
			If FORM_QUESTION_ARRAY(quest).info_type = "utilities" Then call FORM_QUESTION_ARRAY(quest).store_dialog_entry(TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), TEMP_UTILITIES_ARRAY)
			If FORM_QUESTION_ARRAY(quest).info_type = "assets" Then call FORM_QUESTION_ARRAY(quest).store_dialog_entry(TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), TEMP_ASSETS_ARRAY)
			If FORM_QUESTION_ARRAY(quest).info_type = "msa" Then call FORM_QUESTION_ARRAY(quest).store_dialog_entry(TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), TEMP_MSA_ARRAY)
			If FORM_QUESTION_ARRAY(quest).info_type = "stwk" Then call FORM_QUESTION_ARRAY(quest).store_dialog_entry(TEMP_INFO_ARRAY(form_yn_const, quest), TEMP_INFO_ARRAY(form_write_in_const, quest), TEMP_INFO_ARRAY(intv_notes_const, quest), TEMP_INFO_ARRAY(form_second_yn_const, quest), TEMP_STWK_ARRAY)
		End If

		' If FORM_QUESTION_ARRAY(quest).info_type = "housing" Then
		' 	MsgBox "Rent answer: " & FORM_QUESTION_ARRAY(quest).item_ans_list(0)
		' End If
	Next

	For quest = 0 to UBound(FORM_QUESTION_ARRAY)
		If ButtonPressed = FORM_QUESTION_ARRAY(quest).verif_btn Then
			call FORM_QUESTION_ARRAY(quest).capture_verif_detail()
		End If

		If ButtonPressed = FORM_QUESTION_ARRAY(quest).add_to_array_btn Then
			another_job = ""
			count = 0
			for each_item = 0 to UBOUND(FORM_QUESTION_ARRAY(quest).detail_interview_notes)
				count = count + 1
				blank_item = true
				If FORM_QUESTION_ARRAY(quest).detail_source = "jobs" Then
					If FORM_QUESTION_ARRAY(quest).detail_business(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_monthly_amount(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_hourly_wage(each_item) <> "" Then blank_item = false
				ElseIf FORM_QUESTION_ARRAY(quest).detail_source = "assets" Then
					If FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_type(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_value(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_explain(each_item) <> "" Then blank_item = false
				ElseIf FORM_QUESTION_ARRAY(quest).detail_source = "unea" Then
					If FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_type(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_date(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_amount(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_frequency(each_item) <> "" Then blank_item = false
				ElseIf FORM_QUESTION_ARRAY(quest).detail_source = "shel-hest" Then
					If FORM_QUESTION_ARRAY(quest).detail_type(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_amount(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_frequency(each_item) <> "" Then blank_item = false
				ElseIf FORM_QUESTION_ARRAY(quest).detail_source = "expense" Then
					If FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_amount(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_current(each_item) <> "" Then blank_item = false
				ElseIf FORM_QUESTION_ARRAY(quest).detail_source = "winnings" Then
					If FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_amount(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_date(each_item) <> "" Then blank_item = false
				ElseIf FORM_QUESTION_ARRAY(quest).detail_source = "changes" Then
					If FORM_QUESTION_ARRAY(quest).detail_resident_name(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_date(each_item) <> "" Then blank_item = false
					If FORM_QUESTION_ARRAY(quest).detail_explain(each_item) <> "" Then blank_item = false
				End If
				If blank_item = true Then another_job = each_item
			Next
			If another_job = "" Then
				another_job = count
				' MsgBox "another_job - " & another_job & vbCr & IsArray(FORM_QUESTION_ARRAY(quest).detail_interview_notes)
				FORM_QUESTION_ARRAY(quest).add_detail_item(another_job)
				FORM_QUESTION_ARRAY(quest).detail_edit_btn(another_job) = 2000 + quest*10 + another_job

			End If
			Call array_details_dlg(quest, another_job)
		End If
		If IsArray(FORM_QUESTION_ARRAY(quest).detail_interview_notes) = true Then
			for each_item = 0 to UBOUND(FORM_QUESTION_ARRAY(quest).detail_interview_notes)
				If ButtonPressed = FORM_QUESTION_ARRAY(quest).detail_edit_btn(each_item) Then Call array_details_dlg(quest, each_item)
			next
		End If
		' If ButtonPressed = FORM_QUESTION_ARRAY(quest).detail_edit_btn Then
		' End If
	Next

	If ButtonPressed = caf_page_one_btn Then page_display = 1
	If ButtonPressed = caf_addr_btn Then page_display = 2
	If ButtonPressed = caf_membs_btn Then page_display = 3
	If ButtonPressed = caf_q_pg_4 Then page_display = 4
	If ButtonPressed = caf_q_pg_5 Then page_display = 5
	If ButtonPressed = caf_q_pg_6 Then page_display = 6
	If ButtonPressed = caf_q_pg_7 Then page_display = 7
	If ButtonPressed = caf_q_pg_8 Then page_display = 8
	If ButtonPressed = caf_q_pg_9 Then page_display = 9
	If ButtonPressed = caf_q_pg_10 Then page_display = 10
	If ButtonPressed = caf_q_pg_11 Then page_display = 11

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
Set objFrmFld = objDoc.FormFields

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

table_count = 2			'table index variable
ReDim TABLE_ARRAY(0)			'This creates the table array for if there is only one person listed on the CAF
array_counters = 1		'the incrementer for the table array'


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
