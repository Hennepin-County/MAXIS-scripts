IF programs = "Health Care" or programs = "Medical Assistance" THEN CALL create_outlook_email("HSPH.FIN.Unit.AR.Spaulding@hennepin.us", "",
"Claims entered for #" &  MAXIS_case_number, " Member #: " & memb_number & vbcr & " Date Overpayment Created: " & OP_Date & vbcr & "Programs: "
& programs & vbcr & " ", "", False)

Call navigate_to_MAXIS_screen("Case", "note")
EmReadScreen page_line_one, 75, 4, 3
EmReadScreen page_line_two, 75, 5, 3
EmReadScreen page_line_three, 75, 6, 3
EmReadScreen page_line_four, 75, 7, 3
EmReadScreen page_line_five, 75, 8, 3
EmReadScreen page_line_six, 75, 9, 3
EmReadScreen page_line_seven, 75, 10, 3
EmReadScreen page_line_eight, 75, 11, 3
EmReadScreen page_line_nine, 75, 12, 3
EmReadScreen page_line_ten, 75, 13, 3
EmReadScreen page_line_eleven, 75, 14, 3
EmReadScreen page_line_twelve, 75, 15, 3
EmReadScreen page_line_one, 75, 16, 3
EmReadScreen page_line_one, 75, 17, 3

EmReadScreen next_page, 4, 18, 3  If next_page = "more" then
    PF8

    EmReadScreen 75, 4, 3
    EmReadScreen 75, 5, 3
    EmReadScreen 75, 6, 3
    EmReadScreen 75, 7, 3
    EmReadScreen 75, 8, 3
    EmReadScreen 75, 9, 3
    EmReadScreen 75, 10, 3
    EmReadScreen 75, 11, 3
    EmReadScreen 75, 12, 3
    EmReadScreen 75, 13, 3
    EmReadScreen 75, 14, 3
    EmReadScreen 75, 15, 3
    EmReadScreen 75, 16, 3
    EmReadScreen 75, 17, 3



    function write_bullet_and_variable_in_CCOL_NOTE(bullet, variable)
    '--- This function creates an asterisk, a bullet, a colon then a variable to style CCOL notes
    '~~~~~ bullet: name of the field to update. Put bullet in "".
    '~~~~~ variable: variable from script to be written into CCOL note
    '===== Keywords: MAXIS, bullet, CCOL note
    	EMGetCursor noting_row, noting_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
    	noting_col = 3											'The noting col should always be 3 at this point, because it's the beginning. But, this will be dynamically recreated each time.
    	'The following figures out if we need a new page, or if we need a new case note entirely as well.
    	Do
    		EMReadScreen character_test, 1, noting_row, noting_col 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
    		If character_test <> " " or noting_row >= 19 then
    			noting_row = noting_row + 1

    			'If we get to row 18 (which can't be read here), it will go to the next panel (PF8).
    			If noting_row >= 19 then
    				EMSendKey "<PF8>"
    				EMWaitReady 0, 0

    				'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
    				EMReadScreen end_of_case_note_check, 1, 24, 2
    				If end_of_case_note_check = "A" then
    					EMSendKey "<PF3>"												'PF3s
    					EMWaitReady 0, 0
    					EMSendKey "<PF9>"												'PF9s (opens new note)
    					EMWaitReady 0, 0
    					EMWriteScreen "~~~continued from previous note~~~", 5, 	3		'enters a header
    					EMSetCursor 6, 3												'Sets cursor in a good place to start noting.
    					noting_row = 6													'Resets this variable to work in the new locale
    				Else
    					noting_row = 5													'Resets this variable to 5 if we did not need a brand new note.
    				End if
    			End if
    		End if
    	Loop until character_test = " "

    	'Looks at the length of the bullet. This determines the indent for the rest of the info. Going with a maximum indent of 18.
    	If len(bullet) >= 14 then
    		indent_length = 18	'It's four more than the bullet text to account for the asterisk, the colon, and the spaces.
    	Else
    		indent_length = len(bullet) + 4 'It's four more for the reason explained above.
    	End if

    	'Writes the bullet
    	EMWriteScreen "* " & bullet & ": ", noting_row, noting_col

    	'Determines new noting_col based on length of the bullet length (bullet + 4 to account for asterisk, colon, and spaces).
    	noting_col = noting_col + (len(bullet) + 4)

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

    			'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
    			EMReadScreen end_of_case_note_check, 1, 24, 2
    			If end_of_case_note_check = "A" then
    				EMSendKey "<PF3>"												'PF3s
    				EMWaitReady 0, 0
    				EMSendKey "<PF9>"												'PF9s (opens new note)
    				EMWaitReady 0, 0
    				EMWriteScreen "~~~continued from previous note~~~", 4, 	3		'enters a header
    				EMSetCursor 5, 3												'Sets cursor in a good place to start noting.
    				noting_row = 6													'Resets this variable to work in the new locale
    			Else
    				noting_row = 5													'Resets this variable to 4 if we did not need a brand new note.
    			End if
    		End if

    		'Adds spaces (indent) if we're on col 3 since it's the beginning of a line. We also have to increase the noting col in these instances (so it doesn't overwrite the indent).
    		If noting_col = 3 then
    			EMWriteScreen space(indent_length), noting_row, noting_col
    			noting_col = noting_col + indent_length
    		End if

    		'Writes the word and a space using EMWriteScreen
    		EMWriteScreen replace(word, ";", "") & " ", noting_row, noting_col

    		'If a semicolon is seen (we use this to mean "go down a row", it will kick the noting row down by one and add more indent again.
    		If right(word, 1) = ";" then
    			noting_row = noting_row + 1
    			noting_col = 3
    			EMWriteScreen space(indent_length), noting_row, noting_col
    			noting_col = noting_col + indent_length
    		End if

    		'Increases noting_col the length of the word + 1 (for the space)
    		noting_col = noting_col + (len(word) + 1)
    	Next

    	'After the array is processed, set the cursor on the following row, in col 3, so that the user can enter in information here (just like writing by hand). If you're on row 18 (which isn't writeable), hit a PF8. If the panel is at the very end (page 5), it will back out and go into another case note, as we did above.
    	EMSetCursor noting_row + 1, 3
    end function


    function write_variable_in_CCOL_NOTE(variable)
    '--- This function writes a variable in CCOL note
    '~~~~~ variable: information to be entered into CCOL note from script/edit box
    '===== Keywords: MAXIS, CCOL note
    	EMGetCursor noting_row, noting_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
    	noting_col = 3											'The noting col should always be 3 at this point, because it's the beginning. But, this will be dynamically recreated each time.
    	'The following figures out if we need a new page, or if we need a new case note entirely as well.
    	Do
    		EMReadScreen character_test, 1, noting_row, noting_col 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
    		If character_test <> " " or noting_row >= 19 then
    			noting_row = noting_row + 1

    			'If we get to row 19 (which can't be read here), it will go to the next panel (PF8).
    			If noting_row >= 19 then
    				EMSendKey "<PF8>"
    				EMWaitReady 0, 0

    				'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
    				EMReadScreen end_of_case_note_check, 1, 24, 2
    				If end_of_case_note_check = "A" then
    					EMSendKey "<PF3>"												'PF3s
    					EMWaitReady 0, 0
    					EMSendKey "<PF9>"												'PF9s (opens new note)
    					EMWaitReady 0, 0
    					EMWriteScreen "~~~continued from previous note~~~", 4, 	3		'enters a header
    					EMSetCursor 5, 3												'Sets cursor in a good place to start noting.
    					noting_row = 6													'Resets this variable to work in the new locale
    				Else
    					noting_row = 5													'Resets this variable to 5 if we did not need a brand new note.
    				End if
    			End if
    		End if
    	Loop until character_test = " "

    	'Splits the contents of the variable into an array of words
    	variable_array = split(variable, " ")

    	For each word in variable_array

    		'If the length of the word would go past col 80 (you can't write to col 80), it will kick it to the next line and indent the length of the bullet
    		If len(word) + noting_col > 80 then
    			noting_row = noting_row + 1
    			noting_col = 3
    		End if

    		'If the next line is row 19 (you can't write to row 19), it will PF8 to get to the next page
    		If noting_row >= 19 then
    			EMSendKey "<PF8>"
    			EMWaitReady 0, 0

    			'Checks to see if we've reached the end of available case notes. If we are, it will get us to a new case note.
    			EMReadScreen end_of_case_note_check, 1, 24, 2
    			If end_of_case_note_check = "A" then
    				EMSendKey "<PF3>"												'PF3s
    				EMWaitReady 0, 0
    				EMSendKey "<PF9>"												'PF9s (opens new note)
    				EMWaitReady 0, 0
    				EMWriteScreen "~~~continued from previous note~~~", 4, 	3		'enters a header
    				EMSetCursor 5, 3												'Sets cursor in a good place to start noting.
    				noting_row = 6													'Resets this variable to work in the new locale
    			Else
    				noting_row = 5													'Resets this variable to 5 if we did not need a brand new note.
    			End if
    		End if

    		'Writes the word and a space using EMWriteScreen
    		EMWriteScreen replace(word, ";", "") & " ", noting_row, noting_col

    		'Increases noting_col the length of the word + 1 (for the space)
    		noting_col = noting_col + (len(word) + 1)
    	Next

    	'After the array is processed, set the cursor on the following row, in col 3, so that the user can enter in information here (just like writing by hand). If you're on row 18 (which isn't writeable), hit a PF8. If the panel is at the very end (page 5), it will back out and go into another case note, as we did above.
    	EMSetCursor noting_row + 1, 3
    end function


    function start_a_blank_CASE_NOTE()
    '--- This function navigates user to a blank case note, presses PF9, and checks to make sure you're in edit mode (keeping you from writing all of the case note on an inquiry screen).
    '===== Keywords: MAXIS, case note, navigate, edit
    	call navigate_to_MAXIS_screen("case", "note")
    	DO
    		PF9
    		EMReadScreen case_note_check, 17, 2, 33
    		EMReadScreen mode_check, 1, 20, 09
    		If case_note_check <> "Case Notes (NOTE)" or mode_check <> "A" then msgbox "The script can't open a case note. Reasons may include:" & vbnewline & vbnewline & "* You may be in inquiry" & vbnewline & "* You may not have authorization to case note this case (e.g.: out-of-county case)" & vbnewline & vbnewline & "Check MAXIS and/or navigate to CASE/NOTE, and try again. You can press the STOP SCRIPT button on the power pad to stop the script."
    	Loop until (mode_check = "A" or mode_check = "E")
    end function
