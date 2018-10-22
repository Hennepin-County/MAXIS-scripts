'**********THIS IS A HENNEPIN SPECIFIC SCRIPT.  IF YOU REVERSE ENGINEER THIS SCRIPT, JUST BE CAREFUL.************
'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "Make a fucntion.vbs"
start_time = timer
STATS_counter = 0			 'sets the stats counter at one
STATS_manualtime = 60			 'manual run time in seconds
STATS_denomination = "M"		 'M is for Member
'END OF stats block==============================================================================================

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

'FUNCTIONS------------------------------------------------------------------------------------------------
Function get_to_RKEY()
    EMReadScreen MMIS_panel_check, 4, 1, 52	'checking to see if user is on the RKEY panel in MMIS. If not, then it will go to there.
    IF MMIS_panel_check <> "RKEY" THEN
    	DO
    		PF6
    		EMReadScreen session_terminated_check, 18, 1, 7
    	LOOP until session_terminated_check = "SESSION TERMINATED"
    	'Getting back in to MMIS and transmitting past the warning screen (workers should already have accepted the warning screen when they logged themselves into MMIS the first time!)
    	EMWriteScreen "mw00", 1, 2
    	transmit
    	transmit
        EMWriteScreen "x", 9, 4
        transmit
    	EMWriteScreen "x", 8, 3
    	transmit
    END IF
End Function


Function write_variable_in_MMIS_NOTE(variable)
    If trim(variable) <> "" THEN
        EMGetCursor noting_row, noting_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
        noting_col = 8											'The noting col should always be 3 at this point, because it's the beginning. But, this will be dynamically recreated each time.
        'The following figures out if we need a new page, or if we need a new case note entirely as well.
		Do
			EMReadScreen character_test, 1, noting_row, noting_col 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
			If character_test <> " " or noting_row >= 20 then
				noting_row = noting_row + 1

				'If we get to row 18 (which can't be read here), it will go to the next panel (PF8).
				If noting_row >= 20 then
                    PF11
                    noting_row = 5
				End if
			End if
		Loop until character_test = " "

        'Splits the contents of the variable into an array of words
        variable_array = split(variable, " ")

        For each word in variable_array

			'If the length of the word would go past col 80 (you can't write to col 80), it will kick it to the next line and indent the length of the bullet
			If len(word) + noting_col > 72 then
				noting_row = noting_row + 1
				noting_col = 8
			End if

            'If the next line is row 18 (you can't write to row 18), it will PF8 to get to the next page
			If noting_row >= 20 then
                PF11
                noting_row = 5
			End if

            'Writes the word and a space using EMWriteScreen
			EMWriteScreen replace(word, ";", "") & " ", noting_row, noting_col

			'Increases noting_col the length of the word + 1 (for the space)
			noting_col = noting_col + (len(word) + 1)
		Next

		'After the array is processed, set the cursor on the following row, in col 3, so that the user can enter in information here (just like writing by hand). If you're on row 18 (which isn't writeable), hit a PF8. If the panel is at the very end (page 5), it will back out and go into another case note, as we did above.
		EMSetCursor noting_row + 1, 3
    End if
End Function


Function write_bullet_and_variable_in_MMIS_NOTE(bullet, variable)
    If trim(variable) <> "" THEN
        EMGetCursor noting_row, noting_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
        noting_col = 8											'The noting col should always be 3 at this point, because it's the beginning. But, this will be dynamically recreated each time.
        'The following figures out if we need a new page, or if we need a new case note entirely as well.
        Do
            EMReadScreen character_test, 1, noting_row, noting_col 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
            If character_test <> " " or noting_row >= 20 then
                noting_row = noting_row + 1

                'If we get to row 18 (which can't be read here), it will go to the next panel (PF8).
                If noting_row >= 20 then
                    PF11
                    noting_row = 5
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
            If len(word) + noting_col > 72 then
                noting_row = noting_row + 1
                noting_col = 8
            End if

            'If the next line is row 18 (you can't write to row 18), it will PF8 to get to the next page
            If noting_row >= 20 then
                PF11
                noting_row = 5
            End if

            'Adds spaces (indent) if we're on col 3 since it's the beginning of a line. We also have to increase the noting col in these instances (so it doesn't overwrite the indent).
            If noting_col = 8 then
                EMWriteScreen space(indent_length), noting_row, noting_col
                noting_col = noting_col + indent_length
            End if

            'Writes the word and a space using EMWriteScreen
            EMWriteScreen replace(word, ";", "") & " ", noting_row, noting_col

            'If a semicolon is seen (we use this to mean "go down a row", it will kick the noting row down by one and add more indent again.
            If right(word, 1) = ";" then
                noting_row = noting_row + 1
                noting_col = 8
                EMWriteScreen space(indent_length), noting_row, noting_col
                noting_col = noting_col + indent_length
            End if

            'Increases noting_col the length of the word + 1 (for the space)
            noting_col = noting_col + (len(word) + 1)
        Next

        'After the array is processed, set the cursor on the following row, in col 3, so that the user can enter in information here (just like writing by hand). If you're on row 18 (which isn't writeable), hit a PF8. If the panel is at the very end (page 5), it will back out and go into another case note, as we did above.
        EMSetCursor noting_row + 1, 3
    End if
End Function

'SCRIPT--------------------------------------------------------------------------------

EMConnect ""

Call get_to_RKEY

MMIS_case_number = "00269569"

EMWriteScreen "C", 2, 19
EMWriteScreen MMIS_case_number, 9, 19
transmit
transmit
transmit
EMReadscreen RCIN_check, 4, 1, 49
If RCIN_check <> "RCIN" then script_end_procedure("The listed Case number was not found. Check your Case number and try again.")

rcin_row = 11
DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
	EMReadscreen active_status, 1, rcin_row, 76

    If active_status = "A" Then Exit Do

	rcin_row = rcin_row + 1
	If rcin_row = 21 Then
		PF8
		EMReadScreen end_rcin, 6, 24, 2
		If end_rcin = "CANNOT" then Exit Do
		rcin_row = 11
	End If
	Emreadscreen last_clt_check, 8, rcin_row, 4
LOOP until last_clt_check = "        "			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

EMWriteScreen "X", rcin_row, 2
transmit

PF4
PF11

CALL write_variable_in_MMIS_NOTE("*** Start of note ******")

CALL write_bullet_and_variable_in_MMIS_NOTE("LINE 1", "Here are all the things that I need to say - they are important. So much so that I need multiple lines to do all the things that are required")
CALL write_bullet_and_variable_in_MMIS_NOTE("LINE 2", "Here are all the things that I need to say - they are important. So much so that I need multiple lines to do all the things that are required")
CALL write_bullet_and_variable_in_MMIS_NOTE("LINE 3", "Here are all the things that I need to say - they are important. So much so that I need multiple lines to do all the things that are required")
CALL write_bullet_and_variable_in_MMIS_NOTE("LINE 4", "Here are all the things that I need to say - they are important. So much so that I need multiple lines to do all the things that are required")
CALL write_bullet_and_variable_in_MMIS_NOTE("LINE 5", "Here are all the things that I need to say - they are important. So much so that I need multiple lines to do all the things that are required")
CALL write_bullet_and_variable_in_MMIS_NOTE("LINE 6", "Here are all the things that I need to say - they are important. So much so that I need multiple lines to do all the things that are required")
CALL write_bullet_and_variable_in_MMIS_NOTE("LINE 7", "Here are all the things that I need to say; they are important.; So much so that I need; multiple lines to do; all the things that are required")
CALL write_bullet_and_variable_in_MMIS_NOTE("LINE 8", "Here are all the things that I need to say - they are important. So much so that I need multiple lines to do all the things that are required")
CALL write_bullet_and_variable_in_MMIS_NOTE("LINE 9", "Here are all the things that I need to say - they are important. So much so that I need multiple lines to do all the things that are required")


' CALL write_variable_in_MMIS_NOTE("LINE 1")
' CALL write_variable_in_MMIS_NOTE("LINE 2")
' CALL write_variable_in_MMIS_NOTE("LINE 3")
' CALL write_variable_in_MMIS_NOTE("LINE 4")
' CALL write_variable_in_MMIS_NOTE("LINE 5")
' CALL write_variable_in_MMIS_NOTE("LINE 6")
' CALL write_variable_in_MMIS_NOTE("LINE 7")
' CALL write_variable_in_MMIS_NOTE("LINE 8")
' CALL write_variable_in_MMIS_NOTE("LINE 9")
' CALL write_variable_in_MMIS_NOTE("LINE 10")
' CALL write_variable_in_MMIS_NOTE("LINE 11")
' CALL write_variable_in_MMIS_NOTE("LINE 12")
' CALL write_variable_in_MMIS_NOTE("LINE 13")
' CALL write_variable_in_MMIS_NOTE("LINE 14")
' CALL write_variable_in_MMIS_NOTE("LINE 15")
' CALL write_variable_in_MMIS_NOTE("LINE 16")
' CALL write_variable_in_MMIS_NOTE("LINE 17")
' CALL write_variable_in_MMIS_NOTE("LINE 18")
' CALL write_variable_in_MMIS_NOTE("LINE 19")
' CALL write_variable_in_MMIS_NOTE("LINE 20")
' CALL write_variable_in_MMIS_NOTE("LINE 21")
CALL write_variable_in_MMIS_NOTE("C Love - end of note")
