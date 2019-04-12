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


'Going to the MISC panel to add claim referral tracking information
Call navigate_to_MAXIS_screen ("STAT", "MISC")
Row = 6
EmReadScreen panel_number, 1, 02, 78
If panel_number = "0" then
    EMWriteScreen "NN", 20,79
    TRANSMIT
    'CHECKING FOR MAXIS PROGRAMS ARE INACTIVE'
    EmReadScreen MISC_error_msg,  74, 24, 02
    IF trim(MISC_error_msg) = "" THEN
        Do
            'Checking to see if the MISC panel is empty, if not it will find a new line'
            EmReadScreen MISC_description, 25, row, 30
            MISC_description = replace(MISC_description, "_", "")
            If trim(MISC_description) = "" then
                'PF9
                EXIT DO
            Else
                row = row + 1
            End if
        Loop Until row = 17
        If row = 17 then MsgBox("There is not a blank field in the MISC panel. Please delete a line(s), and run script again or update manually.")
        PF9'writing in the action taken and date to the MISC panel
        EMWriteScreen "Claim Determination", Row, 30
        EMWriteScreen date, Row, 66
        PF3
    ELSE
        maxis_error_check = MsgBox("*** NOTICE!!!***" & vbNewLine & "Continue to case note only?" & vbNewLine & MISC_error_msg & vbNewLine, vbYesNo + vbQuestion, "Message handling")
        IF maxis_error_check = vbYes THEN
            case_note_only = TRUE 'this will case note only'
            'should have step to take'
        END IF
        IF maxis_error_check= vbNo THEN
            case_note_only = FALSE 'this will update the panels and case note'
            PF9'writing in the action taken and date to the MISC panel
            EMWriteScreen "Claim Determination", Row, 30
            EMWriteScreen date, Row, 66
            PF3
        END IF
    END IF
ELSE
    Do
        'Checking to see if the MISC panel is empty, if not it will find a new line'
        EmReadScreen MISC_description, 25, row, 30
        MISC_description = replace(MISC_description, "_", "")
        If trim(MISC_description) = "" then
            'PF9
            EXIT DO
        Else
            row = row + 1
        End if
    Loop Until row = 17
    If row = 17 then MsgBox("There is not a blank field in the MISC panel. Please delete a line(s), and run script again or update manually.")
END IF
PF9'writing in the action taken and date to the MISC panel
EMWriteScreen "Claim Determination", Row, 30
EMWriteScreen date, Row, 66
PF3


If trim(variable) <> "" THEN
	EMGetCursor noting_row, noting_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
	noting_col = 3											'The noting col should always be 3 at this point, because it's the beginning. But, this will be dynamically recreated each time.
	'The following figures out if we need a new page, or if we need a new case note entirely as well.
	Do
		EMReadScreen character_test, 40, noting_row, noting_col 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
		character_test = trim(character_test)
		If character_test <> "" or noting_row >= 18 then

			'If we get to row 18 (which can't be read here), it will go to the next panel (PF8).
			If noting_row >= 18 then
				EMSendKey "<PF8>"
				EMWaitReady 0, 0

				EMReadScreen check_we_went_to_next_page, 75, 24, 2
				check_we_went_to_next_page = trim(check_we_went_to_next_page)

				'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
				EMReadScreen end_of_case_note_check, 1, 24, 2
				If end_of_case_note_check = "A" then
					EMSendKey "<PF3>"												'PF3s
					EMWaitReady 0, 0
					EMSendKey "<PF9>"												'PF9s (opens new note)
					EMWaitReady 0, 0
					EMWriteScreen "~~~continued from previous note~~~", 4, 	3		'enters a header
					EMSetCursor 5, 3												'Sets cursor in a good place to start noting.
					noting_row = 5													'Resets this variable to work in the new locale
				ElseIf check_we_went_to_next_page = "PLEASE PRESS PF3 TO EXIT OR FILL PAGE BEFORE SCROLLING TO NEXT PAGE" Then
					noting_row = 4
					Do
						EMReadScreen character_test, 40, noting_row, 3 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
						character_test = trim(character_test)
						If character_test <> "" then noting_row = noting_row + 1
					Loop until character_test = ""
				Else
					noting_row = 4													'Resets this variable to 4 if we did not need a brand new note.
				End If
			Else
				noting_row = noting_row + 1
			End if
		End if
	Loop until character_test = ""

	'Splits the contents of the variable into an array of words
	variable_array = split(variable, " ")

	For each word in variable_array

		'If the length of the word would go past col 80 (you can't write to col 80), it will kick it to the next line and indent the length of the bullet
		If len(word) + noting_col > 80 then
			noting_row = noting_row + 1
			noting_col = 3
		End if

		'If the next line is row 18 (you can't write to row 18), it will PF8 to get to the next page
		If noting_row >= 18 then
			EMSendKey "<PF8>"
			EMWaitReady 0, 0

			EMReadScreen check_we_went_to_next_page, 75, 24, 2
			check_we_went_to_next_page = trim(check_we_went_to_next_page)

			'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
			EMReadScreen end_of_case_note_check, 1, 24, 2
			If end_of_case_note_check = "A" then
				EMSendKey "<PF3>"												'PF3s
				EMWaitReady 0, 0
				EMSendKey "<PF9>"												'PF9s (opens new note)
				EMWaitReady 0, 0
				EMWriteScreen "~~~continued from previous note~~~", 4, 	3		'enters a header
				EMSetCursor 5, 3												'Sets cursor in a good place to start noting.
				noting_row = 5													'Resets this variable to work in the new locale
			ElseIf check_we_went_to_next_page = "PLEASE PRESS PF3 TO EXIT OR FILL PAGE BEFORE SCROLLING TO NEXT PAGE" Then
				noting_row = 4
				Do
					EMReadScreen character_test, 40, noting_row, 3 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
					character_test = trim(character_test)
					If character_test <> "" then noting_row = noting_row + 1
				Loop until character_test = ""
			Else
				noting_row = 4													'Resets this variable to 4 if we did not need a brand new note.
			End if
		End if

		'Writes the word and a space using EMWriteScreen
		EMWriteScreen replace(word, ";", "") & " ", noting_row, noting_col

		'Increases noting_col the length of the word + 1 (for the space)
		noting_col = noting_col + (len(word) + 1)
	Next

	'After the array is processed, set the cursor on the following row, in col 3, so that the user can enter in information here (just like writing by hand). If you're on row 18 (which isn't writeable), hit a PF8. If the panel is at the very end (page 5), it will back out and go into another case note, as we did above.
	EMSetCursor noting_row + 1, 3
End if
end function
