client_not_in_HH = FALSE
If HH_memb <> "01" Then
	Call navigate_to_MAXIS_screen("STAT", "REMO")
	Call write_value_and_transmit(HH_memb, 20, 76)
	EmReadscreen check_if_memb_in_HH, 33, 24, 2
	If check_if_memb_in_HH = "MEMBER " & HH_memb & " IS NOT IN THE HOUSEHOLD" Then
		client_not_in_HH = TRUE
	Else
		EmReadscreen HH_memb_left_date, 8, 8, 53
		EmReadscreen HH_memb_exp_return, 8, 14, 53
		EmReadscreen HH_memb_actual_return, 8, 16, 53

		If HH_memb_left_date <> "__ __ __" and HH_memb_exp_return = "__ __ __" and HH_memb_actual_return = "__ __ __" Then client_not_in_HH = TRUE
	End If
End If

If client_not_in_HH = TRUE Then
	Call navigate_to_MAXIS_screen("STAT", "MEMB")
	EMSetCursor 4, 33
	PF1
	memb_list_row = 9
	Do
		EmReadscreen the_ref_number, 2, memb_list_row, 5
		If the_ref_number = HH_memb Then
			EmReadscreen clt_pmi, 8, memb_list_row, 49
			clt_pmi = trim(clt_pmi)
			Exit Do
		End If

		memb_list_row = memb_list_row + 1
		If memb_list_row = 19 Then
			PF8
			memb_list_row = 9
		End If
	Loop until the_ref_number = "  "

	Call back_to_SELF
	Call Navigate_to_MAXIS_screen("PERS", "    ")

	EmWriteScreen clt_pmi, 15, 36
	transmit

	BeginDialog Dialog1, 0, 0, 216, 135, "Dialog"
	  CheckBox 10, 60, 180, 10, "Check here if client is active SNAP on another case.", new_case_checkbox
	  EditBox 95, 75, 65, 15, new_case_number
	  CheckBox 10, 100, 180, 10, "Check here if client is not active SNAP on any case.", clt_closed_checkbox
	  ButtonGroup ButtonPressed
		OkButton 160, 115, 50, 15
	  Text 10, 10, 205, 10, "It appears that MEMBER 00 has been removed from this case."
	  Text 10, 25, 185, 25, "The script has navigated to a person search for the PMI associated with this member. Review client status and review if client is active SNAP on another case."
	  Text 20, 80, 65, 10, "New case number:"
	EndDialog

	Do
		Do
			err_msg = ""

			dialog Dialog1

			new_case_number = trim(new_case_number)
			If new_case_checkbox = checked AND clt_closed_checkbox = checked Then err_msg = err_msg & vbNewLine & "* Client cannot be both on a new case and completely inactive SNAP. Select one box to check."
			If new_case_checkbox = unchecked AND clt_closed_checkbox = unchecked Then err_msg = err_msg & vbNewLine & "* Indicate if client is still active on SNAP or not. Select one of the boxes."
			If new_case_checkbox = checked AND new_case_number = "" Then err_msg = err_msg & vbNewLine & "* Since the client is active on a new case, the new case number needs to be entered here."

			If err_msg <> "" Then MsgBox "Please Resolve to Continue:" & vbNewLine & err_msg

		Loop until err_msg = ""
		call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
	LOOP UNTIL are_we_passworded_out = false

	If clt_closed_checkbox = checked Then
		BANKED_MONTHS_CASES_ARRAY(remove_case, the_case) = TRUE
		If InStr(BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case), " CLIENT WAS REMOVED FROM THIS CASE AND IS NO LONGER ACTIVE SNAP.") = 0 Then BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case) = BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case) & " CLIENT WAS REMOVED FROM THIS CASE AND IS NO LONGER ACTIVE SNAP."
	End If
	If new_case_checkbox = checked Then
		BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case) = BANKED_MONTHS_CASES_ARRAY(clt_notes, the_case) & " Client moved from SNAP on " & MAXIS_case_number & " to SNAP on case number " & new_case_number & " in the month " & MAXIS_footer_month & "/" & MAXIS_footer_year & "."
		MAXIS_case_number = new_case_number
		BANKED_MONTHS_CASES_ARRAY(case_nbr, the_case) = MAXIS_case_number
		client_not_in_HH = FALSE

		list_row = BANKED_MONTHS_CASES_ARRAY(clt_excel_row, the_case)       'setting the excel row to what was found in the array
		ObjExcel.Cells(list_row, case_nbr_col) = BANKED_MONTHS_CASES_ARRAY(case_nbr, the_case)                              'adding the formatted number to the excel sheet because I am tired of crazy looking excel files

		Call back_to_SELF
		Call navigate_to_MAXIS_screen("STAT", "MEMB")

		EMSetCursor 4, 33
		PF1
		memb_list_row = 9
		Do
			EmReadscreen the_pmi, 8, memb_list_row, 49
			the_pmi = trim(the_pmi)
			If the_pmi = clt_pmi Then
				EmReadscreen ref_nbr, 2, memb_list_row, 5
				Exit Do
			End If

			memb_list_row = memb_list_row + 1
			If memb_list_row = 19 Then
				PF8
				memb_list_row = 9
			End If
		Loop until the_pmi = "  "

		BANKED_MONTHS_CASES_ARRAY(memb_ref_nbr, the_case) = ref_nbr                                                             'resetting this reference number to the one on the new case
		HH_memb = ref_nbr
		ObjExcel.Cells(list_row, memb_nrb_col) = BANKED_MONTHS_CASES_ARRAY(memb_ref_nbr, the_case)                              'adding the formatted number to the excel sheet because I am tired of crazy looking excel files

	End If
End If
