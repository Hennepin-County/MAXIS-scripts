''GATHERING STATS===========================================================================================
name_of_script = "ACTIONS - DEU-MATCH CLEARED.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 300
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================

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

'GATHERING STATS===========================================================================================
name_of_script = "ACTIONS - DEU-MATCH CLEARED CC.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 300
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================
'run_locally = TRUE
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
'FUNCTIONS LIBRARY BLOCK================================================================================================
FUNCTION write_variable_in_CCOL_note_test(variable)
    ''--- This function writes a variable in CCOL note
    '~~~~~ variable: information to be entered into CASE note from script/edit box
    '===== Keywords: MAXIS, CASE note
    If trim(variable) <> "" THEN
    	EMGetCursor noting_row, noting_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
    	'msgbox varible & vbcr & "noting_row " & noting_row
        noting_col = 3											'The noting col should always be 3 at this point, because it's the beginning. But, this will be dynamically recreated each time.
    	'The following figures out if we need a new page, or if we need a new case note entirely as well.
    	Do
    		EMReadScreen character_test, 40, noting_row, noting_col 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
    		character_test = trim(character_test)
    		If character_test <> "" or noting_row >= 19 then
                noting_row = noting_row + 1
    		    'If we get to row 19 (which can't be read here), it will go to the next panel (PF8).
    			If noting_row >= 19 then
    				PF8
                    'msgbox "sent PF8"
    				EMReadScreen next_page_confirmation, 4, 19, 3
                    'msgbox "next_page_confirmation " & next_page_confirmation
    				IF next_page_confirmation = "More" THEN
    					next_page = TRUE
                        noting_row = 5
    				Else
						next_page = FALSE
    				End If
                    'msgbox "next_page " & next_page
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
            If noting_row >= 19 then
                PF8
                noting_row = 5
                'Msgbox "what's Happening? Noting row: " & noting_row
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
    End if
END FUNCTION

function write_bullet_and_variable_in_CCOL_note_test(bullet, variable)
'--- This function creates an asterisk, a bullet, a colon then a variable to style CCOL notes
'~~~~~ bullet: name of the field to update. Put bullet in "".
'~~~~~ variable: variable from script to be written into CCOL note
'===== Keywords: MAXIS, bullet, CCOL note
    If trim(variable) <> "" THEN
        EMGetCursor noting_row, noting_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
        'msgbox varible & vbcr & "noting_row " & noting_row
        noting_col = 3											'The noting col should always be 3 at this point, because it's the beginning. But, this will be dynamically recreated each time.
        'The following figures out if we need a new page, or if we need a new case note entirely as well.
        Do
            EMReadScreen character_test, 40, noting_row, noting_col 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
            character_test = trim(character_test)
            If character_test <> "" or noting_row >= 19 then
                noting_row = noting_row + 1
                'If we get to row 19 (which can't be read here), it will go to the next panel (PF8).
                If noting_row >= 19 then
                    PF8
                    'msgbox "sent PF8"
                    EMReadScreen next_page_confirmation, 4, 19, 3
                    'msgbox "next_page_confirmation " & next_page_confirmation
                    IF next_page_confirmation = "More" THEN
                        next_page = TRUE
                        noting_row = 5
                    Else
                        next_page = FALSE
                    End If
                    'msgbox "next_page " & next_page
                Else
                    noting_row = noting_row + 1
                End if
            End if
        Loop until character_test = ""

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
            If noting_row >= 19 then
                PF8
                noting_row = 5
                'Msgbox "what's Happening? Noting row: " & noting_row
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
    End if
end function
'END FUNCTIONS LIBRARY BLOCK================================================================================================
'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: CALL changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
CALL changelog_update("04/22/2020", "Combined OP script with match cleared, added HH member dialog. Created a new drop down for claim referral tracking.", "MiKayla Handley, Hennepin County")
call changelog_update("08/05/2019", "Updated the term claim referral to use the action taken on MISC as well as to read for active programs.", "MiKayla Handley")
CALL changelog_update("07/17/2019", "Updated script to no longer run off DAIL, it will ask for a case number to ensure all the matches pull correctly.", "MiKayla Handley, Hennepin County")
CALL changelog_update("03/14/2019", "Updated dialog and case note to reflect BE-Child requirements.", "MiKayla Handley, Hennepin County")
CALL changelog_update("04/23/2018", "Updated case note to reflect standard dialog and case note.", "MiKayla Handley, Hennepin County")
CALL changelog_update("02/26/2018", "Merged the claim referral tracking back into the script.", "MiKayla Handley, Hennepin County")
CALL changelog_update("01/16/2018", "Corrected case note for pulling IEVS period.", "MiKayla Handley, Hennepin County")
CALL changelog_update("01/02/2018", "Corrected IEVS match error due to new year.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/27/2017", "Updated to handle clearing the match when the date is over 45 days.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/27/2017", "Updated to handle clearing the match BE-OP entered.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/13/2017", "Updated correct handling for BEER matches.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/08/2017", "Now includes handling for sending the difference notice and clearing the WAGE match including NC codes.", "MiKayla Handley, Hennepin County")
CALL changelog_update("11/27/2017", "Added BP - Wrong Person", "MiKayla Handley, Hennepin County")
CALL changelog_update("11/22/2017", "Updated Non-coop option to the cleared match.", "MiKayla Handley, Hennepin County")
CALL changelog_update("11/21/2017", "Updated to clear match, and added handling for sending the difference notice.", "MiKayla Handley, Hennepin County")
CALL changelog_update("11/14/2017", "Initial version.", "MiKayla Handley, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'---------------------------------------------------------------------THE SCRIPT
testing_run = TRUE
EMConnect ""
CALL MAXIS_case_number_finder (MAXIS_case_number)
'MAXIS_case_number = "2260862"
'---------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 111, 45, "Case Number"
  EditBox 65, 5, 40, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    OkButton 20, 25, 40, 15
    CancelButton 65, 25, 40, 15
  Text 5, 10, 50, 10, "Case Number:"
EndDialog

DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
  		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Enter a valid case number."
  		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)
LOOP UNTIL are_we_passworded_out = false

EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case THEN it gets identified, and will not be updated in MMIS
IF PRIV_check = "PRIV" THEN script_end_procedure("PRIV case, cannot access/update. The script will now end.")

'----------------------------------------------------------------------------------------------------Gathering the member information
CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.

client_array = "Select One:" & "|"

DO								'reads the reference number, last name, first name, and THEN puts it into a single string THEN into the array
EMReadscreen ref_nbr, 3, 4, 33
EMReadScreen access_denied_check, 13, 24, 2
'MsgBox access_denied_check
If access_denied_check = "ACCESS DENIED" Then
	PF10
	last_name = "UNABLE TO FIND"
	first_name = " - Access Denied"
	mid_initial = ""
Else
	EMReadscreen last_name, 25, 6, 30
	EMReadscreen first_name, 12, 6, 63
	EMReadscreen mid_initial, 1, 6, 79
	last_name = trim(replace(last_name, "_", "")) & " "
	first_name = trim(replace(first_name, "_", "")) & " "
	mid_initial = replace(mid_initial, "_", "")
End If
	EMReadscreen MEMB_number, 3, 4, 33
	EMReadscreen last_name, 25, 6, 30
	EMReadscreen first_name, 12, 6, 63
	EMReadscreen mid_initial, 1, 6, 79
    EMReadScreen client_DOB, 10, 8, 42
	EMReadscreen client_SSN, 11, 7, 42
	client_SSN = replace(client_SSN, " ", "")
	last_name = trim(replace(last_name, "_", "")) & " "
	first_name = trim(replace(first_name, "_", "")) & " "
	mid_initial = replace(mid_initial, "_", "")
	client_string = MEMB_number & last_name & first_name & client_SSN
	client_array = client_array & trim(client_string) & "|"

	transmit
	Emreadscreen edit_check, 7, 24, 2
LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

client_array = TRIM(client_array)
client_selection = split(client_array, "|")
CALL convert_array_to_droplist_items(client_selection, hh_member_dropdown)
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 171, 60, "HH Composition"
DropListBox 5, 20, 160, 15, hh_member_dropdown, ievs_member
  ButtonGroup ButtonPressed
    OkButton 70, 40, 45, 15
    CancelButton 120, 40, 45, 15
  Text 5, 5, 165, 10, "Please select the HH Member for the IEVS match:"
EndDialog

DO
    DO
       	err_msg = ""
       	Dialog Dialog1
       	cancel_without_confirmation
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
       LOOP UNTIL err_msg = ""
	CALL check_for_password_without_transmit(are_we_passworded_out)
LOOP UNTIL are_we_passworded_out = false

ievs_member = trim(ievs_member)
IEVS_ssn = right(ievs_member, 9)
IEVS_MEMB_number = left(ievs_member, 2)
'MsgBox IEVS_MEMB_number
CALL navigate_to_MAXIS_screen("INFC" , "____")
CALL write_value_and_transmit("IEVP", 20, 71)
CALL write_value_and_transmit(IEVS_ssn, 3, 63)

EMReadscreen err_msg, 75, 24, 02
err_msg = trim(err_msg)
If err_msg <> "" THEN script_end_procedure_with_error_report("*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine)

'------------------------------------------------------------------selecting the correct wage match
Row = 7
DO
	EMReadScreen IEVS_period, 11, row, 47
	EMReadScreen number_IEVS_type, 3, row, 41
	IF trim(IEVS_period) = "" THEN script_end_procedure_with_error_report("A match for the selected period could not be found. The script will now end.")
	BeginDialog Dialog1, 0, 0, 171, 95, "CASE NUMBER: "  & MAXIS_case_number
  	 Text 5, 10, 100, 10, "Navigate to the correct match:"
  	 Text 5, 25, 150, 10, "Match Type: " & number_IEVS_type
  	 Text 5, 40, 150, 10, "Match Period: "  & IEVS_period
  	 ButtonGroup ButtonPressed
     PushButton 5, 60, 50, 15, "Confirm Match", match_confimation
     PushButton 60, 60, 50, 15, "Next Match", next_match
     PushButton 115, 60, 50, 15, "Next Page", next_page
    CancelButton 60, 80, 50, 15
	EndDialog
	DO
	    DO
	       	err_msg = ""
	       	Dialog Dialog1
			cancel_confirmation
			IF ButtonPressed = next_match THEN
				row = row + 1
				'msgbox "row: " & row
				IF row = 17 THEN
					PF8
					row = 7
					EMReadScreen IEVS_period, 11, row, 47
				END IF
			END IF
			IF ButtonPressed = next_page THEN
				PF8
				row = 7
				EMReadScreen IEVS_period, 11, row, 47
			END IF
			IF ButtonPressed = match_confimation THEN EXIT DO
	        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	       LOOP UNTIL err_msg = ""
		CALL check_for_password_without_transmit(are_we_passworded_out)
	LOOP UNTIL are_we_passworded_out = false
LOOP UNTIL ButtonPressed = match_confimation
'---------------------------------------------------------------------Reading potential errors for out-of-county cases
CALL write_value_and_transmit("U", row, 3)   'navigates to IULA
EMReadScreen OutOfCounty_error, 12, 24, 2
IF OutOfCounty_error = "MATCH IS NOT" THEN
	script_end_procedure_with_error_report("Out-of-county case. Cannot update.")
ELSE
    EMReadScreen number_IEVS_type, 3, 7, 12 'read the match type'
    IF number_IEVS_type = "A30" THEN match_type = "BNDX"
    IF number_IEVS_type = "A40" THEN match_type = "SDXS/I"
    IF number_IEVS_type = "A70" THEN match_type = "BEER"
    IF number_IEVS_type = "A80" THEN match_type = "UNVI"
    IF number_IEVS_type = "A60" THEN match_type = "UBEN"
    IF number_IEVS_type = "A50" or number_IEVS_type = "A51"  THEN match_type = "WAGE"

	IEVS_year = ""
	IF match_type = "WAGE" THEN
		EMReadScreen select_quarter, 1, 8, 14
		EMReadScreen IEVS_year, 4, 8, 22
	ELSEIF match_type = "UBEN" THEN
		EMReadScreen IEVS_month, 2, 5, 68
		EMReadScreen IEVS_year, 4, 8, 71
	ELSEIF match_type = "BEER" THEN
		EMReadScreen IEVS_year, 2, 8, 15
		IEVS_year = "20" & IEVS_year
	ELSEIF match_type = "UNVI" THEN
		EMReadScreen IEVS_year, 4, 8, 15
		'msgbox IEVS_year
		select_quarter = "YEAR"
	END IF
END IF

'--------------------------------------------------------------------Client name
EMReadScreen panel_name, 4, 02, 52
IF panel_name <> "IULA" THEN script_end_procedure_with_error_report("Script did not find IULA.")
EMReadScreen client_name, 35, 5, 24
client_name = trim(client_name)                         'trimming the client name
IF instr(client_name, ",") THEN    						'Most cases have both last name and 1st name. This separates the two names
	length = len(client_name)                           'establishing the length of the variable
	position = InStr(client_name, ",")                  'sets the position at the deliminator (in this case the comma)
	last_name = Left(client_name, position-1)           'establishes client last name as being before the deliminator
	first_name = Right(client_name, length-position)    'establishes client first name as after before the deliminator
ELSEIF instr(first_name, " ") THEN   						'If there is a middle initial in the first name, THEN it removes it
	length = len(first_name)                        	'trimming the 1st name
	position = InStr(first_name, " ")               	'establishing the length of the variable
	first_name = Left(first_name, position-1)       	'trims the middle initial off of the first name
ELSE                                'In cases where the last name takes up the entire space, THEN the client name becomes the last name
	first_name = ""
	last_name = client_name
END IF
first_name = trim(first_name)
IF instr(first_name, " ") THEN   						'If there is a middle initial in the first name, THEN it removes it
	length = len(first_name)                        	'trimming the 1st name
	position = InStr(first_name, " ")               	'establishing the length of the variable
	first_name = Left(first_name, position-1)       	'trims the middle initial off of the first name
END IF

'----------------------------------------------------------------------------------------------------ACTIVE PROGRAMS
EMReadScreen Active_Programs, 13, 6, 68
Active_Programs = trim(Active_Programs)
programs = ""
IF instr(Active_Programs, "D") THEN programs = programs & "DWP, "
IF instr(Active_Programs, "F") THEN programs = programs & "Food Support, "
IF instr(Active_Programs, "H") THEN programs = programs & "Health Care, "
IF instr(Active_Programs, "M") THEN programs = programs & "Medical Assistance, "
IF instr(Active_Programs, "S") THEN programs = programs & "MFIP, "
'trims excess spaces of programs
programs = trim(programs)
'takes the last comma off of programs when autofilled into dialog
IF right(programs, 1) = "," THEN programs = left(programs, len(programs) - 1)
'----------------------------------------------------------------------------------------------------Employer info & difference notice info
IF match_type = "UBEN" THEN income_source = "Unemployment"
IF match_type = "UNVI" THEN income_source = "NON-WAGE"
IF match_type = "WAGE" THEN
	EMReadScreen income_source, 50, 8, 37 'was 37' should be to the right of emplyer and the left of amount
    income_source = trim(income_source)
    length = len(income_source)		'establishing the length of the variable
    'should be to the right of employer and the left of amount '
    IF instr(income_source, " AMOUNT: $") THEN
	    position = InStr(income_source, " AMOUNT: $")    		      'sets the position at the deliminator
	    income_source = Left(income_source, position)  'establishes employer as being before the deliminator
	Elseif instr(income_source, " AMT: $") THEN 					  'establishing the length of the variable
        position = InStr(income_source, " AMT: $")    		      'sets the position at the deliminator
        income_source = Left(income_source, position)  'establishes employer as being before the deliminator
	END IF
END IF
IF match_type = "BEER" THEN
	EMReadScreen income_source, 50, 8, 28 'was 37' should be to the right of emplyer and the left of amount
	income_source = trim(income_source)
	length = len(income_source)		'establishing the length of the variable
	'should be to the right of employer and the left of amount '
    IF instr(income_source, " AMOUNT: $") THEN
	    position = InStr(income_source, " AMOUNT: $")    		      'sets the position at the deliminator
	    income_source = Left(income_source, position)  'establishes employer as being before the deliminator
	Elseif instr(income_source, " AMT: $") THEN 					  'establishing the length of the variable
        position = InStr(income_source, " AMT: $")    		      'sets the position at the deliminator
        income_source = Left(income_source, position)  'establishes employer as being before the deliminator
	END IF
END IF

'----------------------------------------------------------------------------------------------------notice sent
EMReadScreen notice_sent, 1, 14, 37
EMReadScreen sent_date, 8, 14, 68
sent_date = trim(sent_date)

IF sent_date = "" THEN sent_date = "N/A"
IF sent_date <> "" THEN sent_date = replace(sent_date, " ", "/")

EMReadScreen clear_code, 2, 12, 58

'----------------------------------------------------------------Defaulting checkboxes to being checked (per DEU instruction)

IF notice_sent = "N" THEN
    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
	BeginDialog Dialog1, 0, 0, 271, 185, "DIFFERENCE NOTICE NOT SENT FOR: " & MAXIS_case_number
	  DropListBox 85, 90, 70, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", difference_notice_action_dropdown
	  CheckBox 175, 15, 70, 10, "Difference Notice", diff_notice_checkbox
	  CheckBox 175, 25, 90, 10, "Authorization to Release", ATR_verf_checkbox
	  CheckBox 175, 35, 90, 10, "Employment verification", EVF_checkbox
	  CheckBox 175, 45, 80, 10, "Lottery/Gaming Form", lottery_verf_checkbox
	  CheckBox 175, 55, 80, 10, "Rental Income Form", rental_checkbox
	  CheckBox 175, 65, 80, 10, "Other (please specify)", other_checkbox
	  CheckBox 10, 170, 115, 10, "Set a TIKL due to 10 day cutoff", tenday_checkbox
	  DropListBox 145, 120, 115, 15, "Not Needed"+chr(9)+"Initial"+chr(9)+"Overpayment Exists"+chr(9)+"OP Non-Collectible (please specify)"+chr(9)+"No Savings/Overpayment", claim_referral_tracking_dropdown
	  EditBox 50, 145, 215, 15, other_notes
	  Text 5, 10, 165, 10, "Client name: "   & client_name
	  Text 5, 55, 160, 10, "Active Programs: "  & programs
	  Text 5, 70, 165, 15, "Income source:   " & income_source
	  ButtonGroup ButtonPressed
	    OkButton 180, 165, 40, 15
	    CancelButton 225, 165, 40, 15
	  Text 5, 25, 150, 10, "Match Type: " & match_type
	  Text 5, 40, 150, 10, "Match Period: " & IEVS_period
	  GroupBox 170, 5, 95, 75, "Verification(s) Requested: "
	  GroupBox 5, 110, 260, 30, "SNAP or MFIP Federal Food only"
	  Text 10, 125, 130, 10, "Claim Referral Tracking on STAT/MISC:"
	  Text 5, 95, 80, 10, "Send Difference Notice: "
	  Text 5, 150, 40, 10, "Other notes: "
	EndDialog

	DO
    	err_msg = ""
    	Dialog Dialog1
    	cancel_without_confirmation
    	IF difference_notice_action_dropdown = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please select an answer to continue."
    	'IF claim_referral_tracking_dropdown =  "Select One:" and difference_notice_action_dropdown =  "YES" THEN err_msg = err_msg & vbNewLine & "* Please select if the claim referral tracking needs to be updated."
		IF other_checkbox = CHECKED and other_notes = "" THEN err_msg = err_msg & vbNewLine & "* Please ensure you are completing other notes"
    	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password_without_transmit(are_we_passworded_out)
END IF

IF difference_notice_action_dropdown =  "YES" THEN '--------------------------------------------------------------------sending the notice in IULA
    EMwritescreen "005", 12, 46 'writing the resolve time to read for later
    EMwritescreen "Y", 14, 37 'send Notice
	'msgbox "Difference Notice Sent"
	TRANSMIT 'goes into IULA
	'removed the IULB information '
	TRANSMIT'exiting IULA, helps prevent errors when going to the case note
    '-----------------------------------------------------------------------------------Claim Referral Tracking
    action_date = date & ""

ELSEIF notice_sent = "Y" or difference_notice_action_dropdown =  "NO" THEN 'or clear_code <> "__" '
	'-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 326, 170, "MATCH CLEARED - CASE NUMBER: "  & MAXIS_case_number
      EditBox 175, 5, 15, 15, resolve_time
      DropListBox 75, 35, 115, 15, "Select One:"+chr(9)+"CB-Ovrpmt And Future Save"+chr(9)+"CC-Overpayment Only"+chr(9)+"CF-Future Save"+chr(9)+"CA-Excess Assets"+chr(9)+"CI-Benefit Increase"+chr(9)+"CP-Applicant Only Savings"+chr(9)+"BC-Case Closed"+chr(9)+"BE-Child"+chr(9)+"BE-No Change"+chr(9)+"BE-NC-Non-collectible"+chr(9)+"BE-Overpayment Entered"+chr(9)+"BN-Already Known-No Savings"+chr(9)+"BI-Interface Prob"+chr(9)+"BP-Wrong Person"+chr(9)+"BU-Unable To Verify"+chr(9)+"BO-Other"+chr(9)+"NC-Non Cooperation", resolution_status
      DropListBox 120, 50, 70, 15, "Select One:"+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"N/A", change_response
      DropListBox 120, 65, 70, 15, "Select One:"+chr(9)+"DISQ Added"+chr(9)+"DISQ Deleted"+chr(9)+"Pending Verif"+chr(9)+"No"+chr(9)+"N/A", DISQ_action
      EditBox 275, 15, 40, 15, date_received
      CheckBox 200, 30, 70, 10, "Difference Notice", diff_notice_checkbox
      CheckBox 200, 40, 90, 10, "Authorization to Release", ATR_verf_checkBox
      CheckBox 200, 50, 90, 10, "Employment verification", EVF_checkbox
      CheckBox 200, 60, 80, 10, "Lottery/Gaming Form", lottery_verf_checkbox
      CheckBox 200, 70, 80, 10, "Rental Income Form", rental_checkbox
      CheckBox 200, 80, 80, 10, "Other (please specify)", other_checkbox
      EditBox 275, 95, 40, 15, exp_grad_date
      CheckBox 5, 85, 115, 10, "Set a TIKL due to 10 day cutoff", tenday_checkbox
      CheckBox 5, 100, 130, 10, "Overpayment (other programs)", HC_OP_checkbox
      DropListBox 140, 125, 175, 15, "Not Needed"+chr(9)+"Initial"+chr(9)+"Overpayment Exists"+chr(9)+"OP Non-Collectible (please specify)"+chr(9)+"No Savings/Overpayment", claim_referral_tracking_dropdown
      EditBox 50, 150, 180, 15, other_notes
      ButtonGroup ButtonPressed
        OkButton 235, 150, 40, 15
        CancelButton 280, 150, 40, 15
      Text 5, 10, 100, 10, "Match Type: " & match_type
      Text 5, 25, 185, 10, "Match Period: " & IEVS_period
      Text 110, 10, 65, 10, "Resolve time (min): "
      GroupBox 195, 5, 125, 110, "Verification Used to Clear: "
      GroupBox 5, 115, 315, 30, "SNAP or MFIP Federal Food only"
      Text 10, 130, 130, 10, "Claim Referral Tracking on STAT/MISC:"
      Text 5, 40, 60, 10, "Resolution Status: "
      Text 5, 55, 110, 10, "Responded to Difference Notice: "
      Text 5, 70, 75, 10, "DISQ panel addressed:"
      Text 5, 155, 40, 10, "Other notes: "
      Text 200, 20, 75, 10, "Date verif rcvd/on file:"
      Text 200, 100, 65, 10, "Expected grad date:"
    EndDialog

	DO
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
		IF IsNumeric(resolve_time) = false or len(resolve_time) > 3 THEN err_msg = err_msg & vbNewLine & "Please enter a valid numeric resolved time, ie 005."
		IF other_checkbox = CHECKED and other_notes = "" THEN err_msg = err_msg & vbNewLine & "Please advise what other verification was used to clear the match."
		IF change_response = "Select One:" THEN err_msg = err_msg & vbNewLine & "Did the client respond to Difference Notice?"
		IF resolution_status = "Select One:" THEN err_msg = err_msg & vbNewLine & "Please select a resolution status to continue."
		IF resolution_status = "BE-No Change" AND other_notes = "" THEN err_msg = err_msg & vbNewLine & "When clearing using BE other notes must be completed."
		IF resolution_status = "BE-Child" AND exp_grad_date = "" THEN err_msg = err_msg & vbNewLine & "When clearing using BE - Child graduation date and date rcvd must be completed."
		If resolution_status = "CC-Overpayment Only" AND programs = "Health Care" or programs = "Medical Assistance" THEN err_msg = err_msg & vbNewLine & "System does not allow HC or MA cases to be cleared with the code 'CC - Claim Entered'."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password_without_transmit(are_we_passworded_out)

	IF resolution_status = "CC-Overpayment Only" or HC_OP_checkbox = CHECKED THEN
	    discovery_date = date
	    '-------------------------------------------------------------------------------------------------DIALOG
	    Dialog1 = "" 'Blanking out previous dialog detail
	    BeginDialog Dialog1, 0, 0, 361, 260, "MATCH CLEARED - CASE NUMBER: "  & MAXIS_case_number
		  Text 5, 5, 245, 15, "Income source: " & income_source
		  DropListBox 310, 5, 45, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", fraud_referral
	      EditBox 65, 25, 40, 15, discovery_date
	      DropListBox 50, 65, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program
	      EditBox 130, 65, 30, 15, OP_from
	      EditBox 180, 65, 30, 15, OP_to
	      EditBox 245, 65, 35, 15, Claim_number
	      EditBox 305, 65, 45, 15, Claim_amount
	      DropListBox 50, 85, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program_II
	      EditBox 130, 85, 30, 15, OP_from_II
	      EditBox 180, 85, 30, 15, OP_to_II
	      EditBox 245, 85, 35, 15, Claim_number_II
	      EditBox 305, 85, 45, 15, Claim_amount_II
	      DropListBox 50, 105, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program_III
	      EditBox 130, 105, 30, 15, OP_from_III
	      EditBox 180, 105, 30, 15, OP_to_III
	      EditBox 245, 105, 35, 15, claim_number_III
	      EditBox 305, 105, 45, 15, Claim_amount_III
	      DropListBox 50, 125, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program_IV
	      EditBox 130, 125, 30, 15, OP_from_IV
	      EditBox 180, 125, 30, 15, OP_to_IV
	      EditBox 245, 125, 35, 15, claim_number_IV
	      EditBox 305, 125, 45, 15, Claim_amount_IV
		  EditBox 130, 155, 30, 15, HC_from
		  EditBox 180, 155, 30, 15, HC_to
		  EditBox 245, 155, 35, 15, HC_claim_number
		  EditBox 305, 155, 45, 15, HC_claim_amount
		  EditBox 80, 155, 20, 15, HC_resp_memb
		  EditBox 305, 175, 45, 15, Fed_HC_AMT
	      CheckBox 235, 205, 120, 10, "Earned income disregard allowed", EI_checkbox
	      EditBox 70, 200, 160, 15, EVF_used
	      EditBox 200, 25, 45, 15, income_rcvd_date
	      EditBox 70, 220, 285, 15, Reason_OP
		  EditBox 330, 25, 20, 15, OT_resp_memb
		  CheckBox 70, 240, 105, 10, "EVF/ATR is still needed", ATR_needed_checkbox
		  ButtonGroup ButtonPressed
		    OkButton 260, 240, 45, 15
		    CancelButton 310, 240, 45, 15
		  Text 265, 30, 60, 10, "OT resp. Memb #:"
		  Text 260, 10, 50, 10, "Fraud referral:"
		  Text 5, 30, 55, 10, "Discovery date: "
		  Text 5, 205, 65, 10, "Income verif used:"
		  Text 10, 160, 70, 10, "OT resp. Memb(s) #:"
		  Text 230, 180, 75, 10, "Total federal HC AMT:"
		  Text 5, 225, 60, 10, "Reason for Claim:"
		  Text 140, 30, 60, 10, "Date income rcvd: "
		  Text 285, 160, 20, 10, "AMT:"
		  Text 105, 160, 20, 10, "From:"
		  Text 215, 160, 25, 10, "Claim #"
		  Text 165, 160, 10, 10, "To:"
		  GroupBox 5, 145, 350, 50, "HC Programs Only"
		  Text 15, 70, 30, 10, "Program:"
		  Text 165, 70, 10, 10, "To:"
		  GroupBox 5, 45, 350, 100, "Overpayment Information"
		  Text 130, 55, 30, 10, "(MM/YY)"
	      Text 180, 55, 30, 10, "(MM/YY)"
		  Text 15, 70, 30, 10, "Program:"
	      Text 15, 110, 30, 10, "Program:"
	      Text 15, 90, 30, 10, "Program:"
		  Text 15, 130, 30, 10, "Program:"
		  Text 105, 70, 20, 10, "From:"
		  Text 105, 90, 20, 10, "From:"
		  Text 105, 110, 20, 10, "From:"
	      Text 105, 130, 20, 10, "From:"
		  Text 165, 70, 10, 10, "To:"
		  Text 165, 90, 10, 10, "To:"
		  Text 165, 110, 10, 10, "To:"
	      Text 165, 130, 10, 10, "To:"
		  Text 215, 70, 25, 10, "Claim #"
		  Text 215, 90, 25, 10, "Claim #"
		  Text 215, 110, 25, 10, "Claim #"
	      Text 215, 130, 25, 10, "Claim #"
		  Text 285, 70, 20, 10, "AMT:"
		  Text 285, 90, 20, 10, "AMT:"
		  Text 285, 110, 20, 10, "AMT:"
	      Text 285, 130, 20, 10, "AMT:"
		EndDialog
	    Do
	        Do
	        	err_msg = ""
	        	dialog Dialog1
	        	cancel_confirmation
	        	IF fraud_referral = "Select:" THEN err_msg = err_msg & vbnewline & "* You must select a fraud referral entry."
	        	IF trim(Reason_OP) = "" or len(Reason_OP) < 5 THEN err_msg = err_msg & vbnewline & "* You must enter a reason for the overpayment please provide as much detail as possible (min 5)."
	           	IF OP_program_II <> "Select:" THEN
	    			IF OP_from_II = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the month and year overpayment occurred II."
	        		IF Claim_number_II = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the claim number."
	        		IF Claim_amount_II = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the amount of claim."
	        	END IF
	    		IF OP_program_III <> "Select:" THEN
	    			IF OP_from_III = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the month and year overpayment occurred III."
	    			IF Claim_number_III = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the claim number."
	    			IF Claim_amount_III = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the amount of claim."
	    		END IF
	    		IF OP_program_IV <> "Select:" THEN
	    			IF OP_from_IV = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the month and year overpayment occurred IV."
	    			IF Claim_number_IV = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the claim number."
	    			IF Claim_amount_IV = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the amount of claim."
	    		END IF
	        	IF HC_claim_number <> "" THEN
	        		IF HC_from = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the month and year overpayment started."
	        		IF HC_to = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the month and year overpayment ended."
	        		IF HC_claim_amount = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the amount of claim."
	        	END IF
	        	IF EVF_used = "" THEN err_msg = err_msg & vbNewLine & "* Please enter verification used for the income received. If no verification was received enter N/A."
	        	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	        LOOP UNTIL err_msg = ""
	        CALL check_for_password_without_transmit(are_we_passworded_out)
	    Loop until are_we_passworded_out = false
	END IF

	IF resolution_status = "CF-Future Save" THEN
	    '-------------------------------------------------------------------------------------------------DIALOG
	    Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 161, 120, "Cleared CF Future Savings"
  		DropListBox 65, 5, 90, 15, "Select One:"+chr(9)+"Case Became Ineligible"+chr(9)+"Person Removed"+chr(9)+"Benefit Increased"+chr(9)+"Benefit Decreased", IULB_result_dropdown
  		DropListBox 65, 25, 90, 15, "Select One:"+chr(9)+"One Time Only"+chr(9)+"Per Month For Nbr of Months", IULB_method_dropdown
    	EditBox 115, 40, 40, 15, IULB_savings_amount
    	EditBox 125, 60, 15, 15, IULB_start_month
    	EditBox 140, 60, 15, 15, IULB_start_year
    	EditBox 140, 80, 15, 15, IULB_months
    	ButtonGroup ButtonPressed
    	OkButton 60, 100, 45, 15
    	CancelButton 110, 100, 45, 15
    	Text 5, 10, 60, 10, "Results for IULB:"
    	Text 5, 30, 55, 10, "Method for IULB:"
    	Text 55, 45, 55, 10, "Savings Amount:"
    	Text 95, 65, 25, 10, "MM/YY"
    	Text 55, 65, 35, 10, "Start Date:"
    	Text 55, 85, 70, 10, "Months for Method R:"
		EndDialog

	    DO
	    	err_msg = ""
	    	Dialog Dialog1
	    	cancel_confirmation
	    	IF IsNumeric(IULB_savings_amount) = false THEN err_msg = err_msg & vbNewLine & "Please enter a valid numeric amount no decimal."
	    	IF IULB_result_dropdown = "Select One:" THEN err_msg = err_msg & vbNewLine & "Please enter the IULB result."
	    	IF IULB_method_dropdown = "Select One:" THEN err_msg = err_msg & vbNewLine & "Please enter the IULB method."
			IF IULB_result_dropdown <> "Person Removed" and IULB_months <> "" THEN err_msg = err_msg & vbNewLine & "SAVINGS MONTHS NOT ALLOWED WITH MONTH CODE O"
	    	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	    LOOP UNTIL err_msg = ""
	    CALL check_for_password_without_transmit(are_we_passworded_out)

 	    IF IULB_result_dropdown = "Case Became Ineligible" THEN IULB_result = "I"
	    IF IULB_result_dropdown = "Person Removed" THEN IULB_result = "R"
	    IF IULB_result_dropdown = "Benefit Increased" THEN IULB_result = "P"
	    IF IULB_result_dropdown = "Benefit Decreased" THEN IULB_result = "N"
		IF IULB_method_dropdown = "One Time Only" THEN IULB_method = "O"
		IF IULB_method_dropdown = "Per Month For Nbr of Months" THEN IULB_method = "O"
	END IF
END IF
'----------------------------------------------------------------------------------------------------RESOLVING THE MATCH
	EMReadScreen panel_name, 4, 02, 52
	IF panel_name <> "IULA" THEN
		EMReadScreen back_panel_name, 4, 2, 52
		If back_panel_name <> "IEVP" Then
			CALL back_to_SELF
			CALL navigate_to_MAXIS_screen("INFC" , "____")
			CALL write_value_and_transmit("IEVP", 20, 71)
			CALL write_value_and_transmit(client_SSN, 3, 63)
		End If
		CALL write_value_and_transmit("U", row, 3)   'navigates to IULA
	End If

	EMWriteScreen resolve_time, 12, 46	    'resolved notes depending on the resolution_status
	IF resolution_status = "CB-Ovrpmt And Future Save" THEN IULA_res_status = "CB"
	IF resolution_status = "CC-Overpayment Only" THEN IULA_res_status = "CC" 'Claim Entered" CC cannot be used - ACTION CODE FOR ACTH OR ACTM IS INVALID
	IF resolution_status = "CF-Future Save" THEN IULA_res_status = "CF"
	IF resolution_status = "CA-Excess Assets" THEN IULA_res_status = "CA"
	IF resolution_status = "CI-Benefit Increase" THEN IULA_res_status = "CI"
	IF resolution_status = "CP-Applicant Only Savings" THEN IULA_res_status = "CP"
	IF resolution_status = "BC-Case Closed" THEN IULA_res_status = "BC"
	IF resolution_status = "BE-Child" THEN IULA_res_status = "BE"
	IF resolution_status = "BE-No Change" THEN IULA_res_status = "BE"
	IF resolution_status = "BE-Overpayment Entered" THEN IULA_res_status = "BE"
	IF resolution_status = "BE-NC-Non-collectible" THEN IULA_res_status = "BE"
	IF resolution_status = "BI-Interface Prob" THEN IULA_res_status = "BI"
	IF resolution_status = "BN-Already Known-No Savings" THEN IULA_res_status = "BN"
	IF resolution_status = "BP-Wrong Person" THEN IULA_res_status = "BP"
	IF resolution_status = "BU-Unable To Verify" THEN IULA_res_status = "BU"
	IF resolution_status = "BO Other" THEN IULA_res_status = "BO"
	IF resolution_status = "NC-Non Cooperation" THEN IULA_res_status = "NC"

	'checked these all to programS'
	EMwritescreen IULA_res_status, 12, 58
	IF IULA_res_status = "CC" THEN
	    col = 57
	    Do
	    	EMReadscreen action_header, 4, 11, col
	    	If action_header <> "    " Then
	    		If action_header = "ACTH" Then
	    			EMWriteScreen "BE", 12, col+1
	    		Else
	    			EMWriteScreen "CC", 12, col+1
	    		End If
	    	End If
	    		col = col + 6
	    Loop until action_header = "    "
	END IF

	IF change_response = "YES" THEN
		EMwritescreen "Y", 15, 37
	ELSE
		EMwritescreen "N", 15, 37
	END IF

	TRANSMIT 'Going to IULB

	EMReadScreen err_msg, 75, 24, 02
	err_msg = trim(err_msg)
	IF err_msg <> "" THEN
		Dialog1 = "" 'Blanking out previous dialog detail
		  BeginDialog Dialog1, 0, 0, 231, 95, "Maxis Message, please screen shot"
			ButtonGroup ButtonPressed
			OkButton 135, 75, 45, 15
			CancelButton 180, 75, 45, 15
			GroupBox 5, 0, 220, 50, "You can update maxis if there is an error, then hit ok to continue."
			Text 15, 10, 190, 35, err_msg
			EditBox 50, 55, 175, 15, email_BZST
			Text 5, 60, 45, 10, "Email BZST:"
		  EndDialog

		'Showing case number dialog
		Do
		  Dialog Dialog1
		  cancel_without_confirmation
		  CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
		Loop until are_we_passworded_out = false					'loops until user passwords back in
		IF email_BZST <> "" THEN CALL create_outlook_email("Mikayla.Handley@hennepin.us", "", "Case #" & maxis_case_number & " Error message: " & err_msg & "  EOM.", "", "", TRUE)
	END IF

    '----------------------------------------------------------------------------------------writing the note on IULB

	IF resolution_status = "CB-Ovrpmt And Future Save" THEN EMWriteScreen "OP Claim entered and future savings." & other_notes, 8, 6
	IF resolution_status = "CC-Overpayment Only" Or HC_OP_checkbox = CHECKED THEN
		EMWriteScreen "OP Claim entered." & other_notes, 8, 6
		CALL clear_line_of_text(8, 6)
		EMWriteScreen "Claim entered. See Case Note. ", 8, 6
		CALL clear_line_of_text(17, 9)
		If action_header <> "ACTH" THEN
			EMWriteScreen Claim_number, 17, 9
			EMWriteScreen Claim_number_II, 18, 9
			EMWriteScreen claim_number_III, 19, 9
		END IF
		'need to check about adding for multiple claims'

		TRANSMIT 'this will take us back to IEVP main menu'

		EMReadScreen err_msg, 75, 24, 02
		err_msg = trim(err_msg)
		IF err_msg <> "" THEN
			Dialog1 = "" 'Blanking out previous dialog detail
			  BeginDialog Dialog1, 0, 0, 231, 95, "Maxis Message, please screen shot"
				ButtonGroup ButtonPressed
				OkButton 135, 75, 45, 15
				CancelButton 180, 75, 45, 15
				GroupBox 5, 0, 220, 50, "You can update maxis if there is an error, THEN hit ok to continue."
				Text 15, 10, 190, 35, err_msg
				EditBox 50, 55, 175, 15, email_BZST
				Text 5, 60, 45, 10, "Email BZST:"
			  EndDialog

			'Showing case number dialog
			Do
			  Dialog Dialog1
			  cancel_without_confirmation
			  CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
			Loop until are_we_passworded_out = false					'loops until user passwords back in
			IF email_BZST <> "" THEN CALL create_outlook_email("Mikayla.Handley@hennepin.us", "", "Case #" & maxis_case_number & " Error message: " & err_msg & "  EOM.", "", "", TRUE)
		END IF
	END IF

	IF resolution_status = "CF-Future Save" THEN
		EMWriteScreen "Future Savings. " & other_notes, 8, 6
		EMwritescreen active_programs, 12, 37
		EMwritescreen IULB_results, 12, 42
		EMwritescreen IULB_method, 12, 49
		EMwritescreen IULB_savings_amount, 12, 54
		EMwritescreen IULB_start_month, 12, 65
		EMwritescreen IULB_start_year, 12, 68
		EMwritescreen IULB_months, 12, 74
		TRANSMIT
	END IF

	IF resolution_status = "CA-Excess Assets" THEN EMWriteScreen "Excess Assets. " & other_notes, 8, 6
	IF resolution_status = "CI-Benefit Increase" THEN EMWriteScreen "Benefit Increase. " & other_notes, 8, 6
	IF resolution_status = "CP-Applicant Only Savings" THEN EMWriteScreen "Applicant Only Savings. " & other_notes, 8, 6
	IF resolution_status = "BC-Case Closed" THEN EMWriteScreen "Case closed. " & other_notes, 8, 6
	IF resolution_status = "BE-Child" THEN EMWriteScreen "No change, minor child income excluded. " & other_notes, 8, 6
	IF resolution_status = "BE-No Change" THEN EMWriteScreen "No change. " & other_notes, 8, 6
	IF resolution_status = "BE-Overpayment Entered" THEN EMWriteScreen "OP entered other programs. " & other_notes, 8, 6
	IF resolution_status = "BE-NC-Non-collectible" THEN EMWriteScreen "Non-Coop remains, but claim is non-collectible. ", 8, 6
	IF resolution_status = "BI-Interface Prob" THEN EMWriteScreen "Interface Problem. " & other_notes, 8, 6
	IF resolution_status = "BN-Already Known-No Savings" THEN EMWriteScreen "Already known - No savings. " & other_notes, 8, 6
	IF resolution_status = "BP-Wrong Person" THEN EMWriteScreen "Client name and wage earner name are different. " & other_notes, 8, 6
	IF resolution_status = "BU-Unable To Verify" THEN EMWriteScreen "Unable To Verify. " & other_notes, 8, 6
	IF resolution_status = "BO Other" THEN EMWriteScreen "HC Claim entered. " & other_notes, 8, 6
	IF resolution_status = "NC-Non Cooperation" THEN EMWriteScreen "Non-coop, requested verf not in ECF, " & other_notes, 8, 6

	'------------------------------------------------------------------back on the IEVP menu, making sure that the match cleared

	'EMReadScreen days_pending, 5, row, 72
	'days_pending = trim(days_pending)
	'match_cleared = TRUE
	'IF IsNumeric(days_pending) = TRUE THEN match_cleared = FALSE
	'If match_cleared = FALSE and sent_date <> date THEN
	'   	confirm_cleared = MsgBox ("The script cannot identify that this match has cleared." & vbNewLine & vbNewLine & "Review IEVP and find the match that is being cleared with this run." &vbNewLine & " ** HAS THE MATCH BEEN CLEARED? **", vbQuestion + vbYesNo, "Confirm Match Cleared")
	'   	IF confirm_cleared = vbYes Then match_cleared = TRUE
	'	IF confirm_cleared = vbno Then
	'		match_cleared = FALSE
	'		script_end_procedure_with_error_report("This match did not clear in IEVP, please advise what may have happened.")
	'	END IF
	'End If
	'--------------------------------------------------------------------The case note & case note related code
	verifcation_needed = ""
  	IF Diff_Notice_Checkbox = CHECKED THEN verifcation_needed = verifcation_needed & "Difference Notice, "
	IF EVF_checkbox = CHECKED THEN verifcation_needed = verifcation_needed & "EVF, "
	IF ATR_Verf_CheckBox = CHECKED THEN verifcation_needed = verifcation_needed & "ATR, "
	IF lottery_verf_checkbox = CHECKED THEN verifcation_needed = verifcation_needed & "Lottery/Gaming Form, "
	IF rental_checkbox =  CHECKED THEN verifcation_needed = verifcation_needed & "Rental Income Form, "
	IF other_checkbox = CHECKED THEN verifcation_needed = verifcation_needed & "Other, "

	IF MAXIS_error_message <> "" THEN
		EMReadScreen MAXIS_error_message, 75, 24, 02
		MAXIS_error_message = trim(MAXIS_error_message)

		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 231, 95, "Maxis Message, please screen shot"
			ButtonGroup ButtonPressed
			OkButton 135, 75, 45, 15
			CancelButton 180, 75, 45, 15
			GroupBox 5, 0, 220, 50, "You can update maxis if there is an error, THEN hit ok to continue."
			Text 15, 10, 190, 35, MAXIS_error_message
			EditBox 50, 55, 175, 15, email_BZST
			Text 5, 60, 45, 10, "Email BZST:"
		EndDialog

		'Showing case number dialog
		Do
		  Dialog Dialog1
		  cancel_without_confirmation
		  CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
		Loop until are_we_passworded_out = false					'loops until user passwords back in
		IF email_BZST <> "" THEN CALL create_outlook_email("Mikayla.Handley@hennepin.us", "", "Case #" & maxis_case_number & " Error message: " & MAXIS_error_message & "  EOM.", "", "", TRUE)
	END IF
'------------------------------------------------------------------STAT/MISC for claim referral tracking
	IF claim_referral_tracking_dropdown <> "Not Needed" THEN
	    'Going to the MISC panel to add claim referral tracking information
		CALL navigate_to_MAXIS_screen ("STAT", "MISC")
		Row = 6
		EMReadScreen panel_number, 1, 02, 73
		If panel_number = "0" THEN
			EMWriteScreen "NN", 20,79
			TRANSMIT
			'CHECKING FOR MAXIS PROGRAMS ARE INACTIVE'
			EMReadScreen MISC_error_check,  74, 24, 02
			IF trim(MISC_error_check) = "" THEN
				case_note_only = FALSE
			else
				maxis_error_check = MsgBox("*** NOTICE!!!***" & vbNewLine & "Continue to case note only?" & vbNewLine & MISC_error_check & vbNewLine, vbYesNo + vbQuestion, "Message handling")
				IF maxis_error_check = vbYes THEN
					case_note_only = TRUE 'this will case note only'
				END IF
				IF maxis_error_check= vbNo THEN
					case_note_only = FALSE 'this will update the panels and case note'
				END IF
			END IF
		END IF

		Do
			'Checking to see if the MISC panel is empty, if not it will find a new line'
			EMReadScreen MISC_description, 25, row, 30
			MISC_description = replace(MISC_description, "_", "")
			If trim(MISC_description) = "" THEN
				'PF9
				EXIT DO
			Else
				row = row + 1
			End if
		Loop Until row = 17
		If row = 17 THEN MsgBox("There is not a blank field in the MISC panel. Please delete a line(s), and run script again or update manually.")

		'writing in the action taken and date to the MISC panel
		PF9
		'_________________________ 25 characters to write on MISC
		IF claim_referral_tracking_dropdown =  "Initial" THEN MISC_action_taken = "Claim Referral Initial"
		IF claim_referral_tracking_dropdown =  "OP Non-Collectible (please specify)" THEN MISC_action_taken = "Determination-Non-Collect"
		IF claim_referral_tracking_dropdown =  "No Savings/Overpayment" THEN MISC_action_taken = "Determination-No Savings"
		IF claim_referral_tracking_dropdown =  "Overpayment Exists" THEN MISC_action_taken =  "Determination-OP Entered" '"Claim Determination 25 character available
		EMWriteScreen MISC_action_taken, Row, 30
		EMWriteScreen date, Row, 66
        TRANSMIT
	END IF
	'------------------------------------------setting up case note header'
	IF ATR_needed_checkbox = CHECKED THEN
		header_note = "ATR/EVF STILL REQUIRED"
	ELSEIF difference_notice_action_dropdown = "YES" THEN
		cleared_header = "DIFF NOTICE SENT"
		sent_date = date
	ELSEIF resolution_status = "CC-Overpayment Only" or HC_OP_checkbox = CHECKED THEN
		cleared_header = "CLEARED CLAIM ENTERED "
	ELSEIF resolution_status = "NC-Non Cooperation" THEN
			cleared_header = "NON-COOPERATION "
	ELSEIF resolution_status <> "CC-Overpayment Only" OR resolution_status <> "NC-Non Cooperation" THEN
		cleared_header = "CLEARED " & IULA_res_status
	ELSEIF resolution_status = "BE-NC-Non-collectible" THEN
		cleared_header = "CLEARED " & IULA_res_status & "Non-Collectible"
	END IF

	IF match_type = "BEER" THEN match_type_letter = "B"
	IF match_type = "UBEN" THEN match_type_letter = "U"
	IF match_type = "UNVI" THEN match_type_letter = "U"

	verifcation_needed = trim(verifcation_needed) 	'takes the last comma off of verifcation_needed when autofilled into dialog if more more than one app date is found and additional app is selected
	IF right(verifcation_needed, 1) = "," THEN verifcation_needed = left(verifcation_needed, len(verifcation_needed) - 1)
	IF match_type = "WAGE" THEN
		IF select_quarter = 1 THEN IEVS_quarter = "1ST"
		IF select_quarter = 2 THEN IEVS_quarter = "2ND"
		IF select_quarter = 3 THEN IEVS_quarter = "3RD"
		IF select_quarter = 4 THEN IEVS_quarter = "4TH"
	END IF

	IEVS_period = trim(IEVS_period)
	IF match_type <> "UBEN" THEN IEVS_period = replace(IEVS_period, "/", " to ")
	IF match_type = "UBEN" THEN IEVS_period = replace(IEVS_period, "-", "/")
	Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days

	'-------------------------------------------------------------------------------------------------The case note
	IF claim_referral_tracking_dropdown <> "Not Needed" THEN
	    start_a_blank_case_note
	    IF claim_referral_tracking_dropdown =  "Initial" THEN
			CALL write_variable_in_case_note("Claim Referral Tracking - Initial")
		ELSE
			CALL write_variable_in_case_note("Claim Referral Tracking - " & MISC_action_taken)
		END IF
	    CALL write_bullet_and_variable_in_case_note("Action Date", action_date)
	    CALL write_bullet_and_variable_in_case_note("Active Program(s)", programs)
	    CALL write_bullet_and_variable_in_case_note("Other Notes", other_notes)
	    CALL write_variable_in_case_note("* Entries for these potential claims must be retained until further notice.")
	    IF case_note_only = TRUE THEN CALL write_variable_in_case_note("Maxis case is inactive unable to add or update MISC panel")
	    CALL write_variable_in_case_note("-----")
	    CALL write_variable_in_case_note(worker_signature)
	    PF3
	END IF
	start_a_blank_case_note
	IF match_type = "WAGE" THEN CALL write_variable_in_case_note("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH"  & " (" & first_name & ") " & cleared_header & header_note & "-----")
	IF match_type = "BEER" THEN CALL write_variable_in_case_note("-----" & IEVS_year & " NON-WAGE MATCH(" & match_type_letter & ")" & " (" & first_name & ") " & cleared_header & header_note & "-----")
	IF match_type = "UNVI" THEN CALL write_variable_in_case_note("-----" & IEVS_year & " NON-WAGE MATCH(" & match_type_letter & ")" & " (" & first_name & ") " & cleared_header & header_note & "-----")
	IF match_type = "UBEN" THEN CALL write_variable_in_case_note("-----" & IEVS_period & " NON-WAGE MATCH(" & match_type_letter & ")" & " (" & first_name & ") " & cleared_header & header_note & "-----")
	CALL write_bullet_and_variable_in_case_note("Discovery date", discovery_date)
	CALL write_bullet_and_variable_in_case_note("Period", IEVS_period)
	CALL write_bullet_and_variable_in_case_note("Active Programs", programs)
	CALL write_bullet_and_variable_in_case_note("Source of income", income_source)
	CALL write_variable_in_case_note("----- ----- ----- ----- ----- ----- -----")
	CALL write_bullet_and_variable_in_case_note("Date Diff notice sent", sent_date)
	IF  difference_notice_action_dropdown = "YES" THEN 
		CALL write_bullet_and_variable_in_case_note("Verifications Requested", verifcation_needed)
		CALL write_variable_in_case_note("* Client must be provided 10 days to return requested verifications")
	ELSE
		CALL write_bullet_and_variable_in_case_note("Verifications Received", verifcation_needed)
	END IF
	IF change_response <> "N/A" THEN CALL write_bullet_and_variable_in_case_note("Responded to Difference Notice", change_response)
	IF DISQ_action <> "Select One:" THEN CALL write_bullet_and_variable_in_case_note("STAT/DISQ addressed for each program", DISQ_action)
	CALL write_bullet_and_variable_in_case_note("Date verification received in ECF", date_received)
	IF resolution_status = "CB-Ovrpmt And Future Save" THEN CALL write_variable_in_case_note("* OP Claim entered and future savings.")
	IF resolution_status = "CF-Future Save" THEN CALL write_variable_in_case_note("* Future Savings.")
	IF resolution_status = "CA-Excess Assets" THEN CALL write_variable_in_case_note("* Excess Assets.")
	IF resolution_status = "CI-Benefit Increase" THEN CALL write_variable_in_case_note("* Benefit Increase.")
	IF resolution_status = "CP-Applicant Only Savings" THEN CALL write_variable_in_case_note("* Applicant Only Savings.")
	IF resolution_status = "BC-Case Closed" THEN CALL write_variable_in_case_note("* Case closed.")
	IF resolution_status = "BE-Child" THEN
		CALL write_variable_in_case_note("* Income is excluded for minor child in school.")
		CALL write_bullet_and_variable_in_case_note("Expected graduation date", exp_grad_date)
	END IF
	IF resolution_status = "BE-No Change" THEN CALL write_variable_in_case_note("* No Overpayments or savings were found related to this match.")
	IF resolution_status = "BE-Overpayment Entered" THEN CALL write_variable_in_case_note("* Overpayments or savings were found related to this match.")
	IF resolution_status = "BE-NC-Non-collectible" THEN CALL write_variable_in_case_note("* No collectible overpayments or savings were found related to this match. Client is still non-coop.")
	IF resolution_status = "BI-Interface Prob" THEN CALL write_variable_in_case_note("* Interface Problem.")
	IF resolution_status = "BN-Already Known-No Savings" THEN CALL write_variable_in_case_note("* Client reported income. Correct income is in JOBS/BUSI and budgeted.")
	IF resolution_status = "BP-Wrong Person" THEN CALL write_variable_in_case_note("* Client name and wage earner name are different.  Client's SSN has been verified. No overpayment or savings related to this match.")
	IF resolution_status = "BU-Unable To Verify" THEN CALL write_variable_in_case_note("* Unable to verify, due to:")
	IF resolution_status = "BO Other" THEN CALL write_variable_in_case_note("* HC Claim entered.")
	IF resolution_status = "NC-Non Cooperation" THEN
		CALL write_variable_in_case_note("* Client failed to cooperate wth wage match.")
		CALL write_variable_in_case_note("* Case approved to close.")
		CALL write_variable_in_case_note("* Client needs to provide: ATR, Income Verification, Difference Notice.")
	END IF
	IF resolution_status = "CC-Overpayment Only" or HC_OP_checkbox = CHECKED THEN
	    CALL write_variable_in_case_note(OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number & " Amt $" & Claim_amount)
	    IF OP_program_II <> "Select:" THEN CALL write_variable_in_case_note(OP_program_II & " Overpayment " & OP_from_II & " through " & OP_to_II & " Claim #" & Claim_number_II & " Amt $" & Claim_amount_II)
	    IF OP_program_III <> "Select:" THEN CALL write_variable_in_case_note(OP_program_III & " Overpayment " & OP_from_III & " through " & OP_to_III & " Claim #" & Claim_number_III & " Amt $" & Claim_amount_III)
	    IF OP_program_IV <> "Select:" THEN CALL write_variable_in_case_note(OP_program_IV & " Overpayment " & OP_from_IV & " through " & OP_to_IV & " Claim #" & Claim_number_IV & " Amt $" & Claim_amount_IV)
	    IF HC_claim_number <> "" THEN
	    	CALL write_variable_in_case_note("HC OVERPAYMENT " & HC_from & " through " & HC_to & " Claim #" & HC_claim_number & " Amt $" & HC_Claim_amount)
	    	CALL write_bullet_and_variable_in_case_note("Health Care responsible members", HC_resp_memb)
	    	CALL write_bullet_and_variable_in_case_note("Total Federal Health Care amount", Fed_HC_AMT)
	    	CALL write_variable_in_case_note("* Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
	    END IF
	    IF EI_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Earned Income Disregard Allowed")
	    IF EI_checkbox = UNCHECKED THEN CALL write_variable_in_case_note("* Earned Income Disregard Not Allowed")
	    CALL write_bullet_and_variable_in_case_note("Fraud referral made", fraud_referral)
	    CALL write_bullet_and_variable_in_case_note("Income verification received", EVF_used)
	    CALL write_bullet_and_variable_in_case_note("Date verification received", income_rcvd_date)
	    CALL write_bullet_and_variable_in_case_note("Reason for overpayment", Reason_OP)
	    CALL write_bullet_and_variable_in_case_note("Other responsible member(s)", OT_resp_memb)
	END IF
	CALL write_bullet_and_variable_in_case_note("Other Notes", other_notes)
	CALL write_variable_in_case_note("----- ----- ----- ----- ----- ----- -----")
	CALL write_variable_in_case_note("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1")
	PF3 'to save casenote'

   	IF resolution_status = "CC-Overpayment Only" or HC_OP_checkbox = CHECKED THEN '-----------------------------------------------------------------------------------------OP CASENOTE
	    IF HC_claim_number <> "" THEN
	    	EMWriteScreen "x", 5, 3
	    	TRANSMIT
	    	note_row = 4			'Beginning of the case notes
	    	Do 						'Read each line
	    		EMReadScreen note_line, 76, note_row, 3
	    		note_line = trim(note_line)
	    		If trim(note_line) = "" Then Exit Do		'Any blank line indicates the end of the case note because there can be no blank lines in a note
	    		message_array = message_array & note_line & vbcr		'putting the lines together
	    		note_row = note_row + 1
	    		If note_row = 18 THEN 									'End of a single page of the case note
	    			EMReadScreen next_page, 7, note_row, 3
	    			If next_page = "More: +" Then 						'This indicates there is another page of the case note
	    				PF8												'goes to the next line and resets the row to read'\
	    				note_row = 4
	    			End If
	    		End If
	    	Loop until next_page = "More:  " OR next_page = "       "	'No more pages
	    	'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
	    	CALL create_outlook_email("HSPH.FIN.Unit.AR.Spaulding@hennepin.us", "","Claims entered for #" &  MAXIS_case_number & " Member # " & memb_number & " Date Overpayment Created: " & discovery_date & "HC Claim # " & HC_claim_number, "CASE NOTE" & vbcr & message_array,"", False)
	    END IF
'-----------------------------------------------------------------writing the CCOL case note'
	    'msgbox "Navigating to CCOL to add case note, please contact the BlueZone Scripts team with any concerns."
	    CALL navigate_to_MAXIS_screen("CCOL", "CLSM")
	    EMWriteScreen Claim_number, 4, 9
	    TRANSMIT
	    'checking for error messages'
		IF MAXIS_error_message <> "" THEN
			EMReadScreen MAXIS_error_message, 75, 24, 02
			MAXIS_error_message = trim(MAXIS_error_message)

			Dialog1 = "" 'Blanking out previous dialog detail
			BeginDialog Dialog1, 0, 0, 231, 95, "Maxis Message, please screen shot"
				ButtonGroup ButtonPressed
				OkButton 135, 75, 45, 15
				CancelButton 180, 75, 45, 15
				GroupBox 5, 0, 220, 50, "You can update maxis if there is an error, THEN hit ok to continue."
				Text 15, 10, 190, 35, MAXIS_error_message
				EditBox 50, 55, 175, 15, email_BZST
				Text 5, 60, 45, 10, "Email BZST:"
			EndDialog

			'Showing case number dialog
			Do
			  Dialog Dialog1
			  cancel_without_confirmation
			  CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
			Loop until are_we_passworded_out = false					'loops until user passwords back in
			IF email_BZST <> "" THEN CALL create_outlook_email("Mikayla.Handley@hennepin.us", "", "Case #" & maxis_case_number & " Error message: " & MAXIS_error_message & "  EOM.", "", "", TRUE)
		END IF
	    PF4
	    EMReadScreen existing_case_note, 1, 5, 6
	    IF existing_case_note = "" THEN
	    	PF4
	    ELSE
	    	PF9
	    END IF

	    IF match_type = "WAGE" THEN CALL write_variable_in_CCOL_note_test("-----" & IEVS_quarter & " QTR " & IEVS_year & "WAGE MATCH"  & " (" & first_name & ") CLEARED CC-CLAIM ENTERED " & header_note & "-----")
        IF match_type = "BEER" or match_type = "UNVI" THEN CALL write_variable_in_CCOL_note_test("-----" & IEVS_year & " NON-WAGE MATCH(" & match_type_letter & ") " & " (" & first_name & ") CLEARED CC-CLAIM ENTERED " & header_note & "-----")
        IF match_type = "UBEN" THEN CALL write_variable_in_CCOL_note_test("-----" & IEVS_period & " NON-WAGE MATCH(" & match_type_letter & ") " & " (" & first_name & ") CLEARED CC-CLAIM ENTERED " & header_note & "-----")
	    CALL write_bullet_and_variable_in_CCOL_NOTE_test("Discovery date", discovery_date)
	    CALL write_bullet_and_variable_in_CCOL_NOTE_test("Period", IEVS_period)
	    CALL write_bullet_and_variable_in_CCOL_NOTE_test("Active Programs", programs)
	    CALL write_bullet_and_variable_in_CCOL_NOTE_test("Source of income", income_source)
	    CALL write_variable_in_CCOL_note_test("----- ----- ----- ----- ----- ----- -----")
        CALL write_variable_in_CCOL_note_test(OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number & " Amt $" & Claim_amount)
        IF OP_program_II <> "Select:" THEN CALL write_variable_in_CCOL_note_test(OP_program_II & " Overpayment " & OP_from_II & " through " & OP_to_II & " Claim #" & Claim_number_II & " Amt $" & Claim_amount_II)
        IF OP_program_III <> "Select:" THEN CALL write_variable_in_CCOL_note_test(OP_program_III & " Overpayment " & OP_from_III & " through " & OP_to_III & " Claim #" & Claim_number_III & " Amt $" & Claim_amount_III)
        IF OP_program_IV <> "Select:" THEN CALL write_variable_in_CCOL_note_test(OP_program_IV & " Overpayment " & OP_from_IV & " through " & OP_to_IV & " Claim #" & Claim_number_IV & " Amt $" & Claim_amount_IV)
        IF HC_claim_number <> "" THEN
        	CALL write_variable_in_CCOL_note_test("HC OVERPAYMENT " & HC_from & " through " & HC_to & " Claim #" & HC_claim_number & " Amt $" & HC_Claim_amount)
        	CALL write_bullet_and_variable_in_CCOL_NOTE_test("Health Care responsible members", HC_resp_memb)
        	CALL write_bullet_and_variable_in_CCOL_NOTE_test("Total Federal Health Care amount", Fed_HC_AMT)
        	CALL write_variable_in_CCOL_note_test("* Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
        END IF
	    IF EI_checkbox = CHECKED THEN CALL write_variable_in_CCOL_note_test("* Earned Income Disregard Allowed")
        IF EI_checkbox = UNCHECKED THEN CALL write_variable_in_CCOL_note_test("* Earned Income Disregard Not Allowed")
        CALL write_bullet_and_variable_in_CCOL_NOTE_test("Fraud referral made", fraud_referral)
        CALL write_bullet_and_variable_in_CCOL_NOTE_test("Income verification received", EVF_used)
        CALL write_bullet_and_variable_in_CCOL_NOTE_test("Date verification received", income_rcvd_date)
        CALL write_bullet_and_variable_in_CCOL_NOTE_test("Reason for overpayment", Reason_OP)
        CALL write_bullet_and_variable_in_CCOL_NOTE_test("Other responsible member(s)", OT_resp_memb)
        CALL write_variable_in_CCOL_note_test("----- ----- ----- ----- ----- ----- -----")
        CALL write_variable_in_CCOL_note_test("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1")
        PF3 'to save CCOL casenote'
	    'script_end_procedure_with_error_report("Overpayment case note entered and copied to CCOL, please review the case to make sure the notes updated correctly." & vbcr & next_page)

		'-------------------------------The following will generate a TIKL formatted date for 10 days from now, and add it to the TIKL
		IF tenday_checkbox = CHECKED THEN CALL create_TIKL("Unable to close due to 10 day cutoff. Verification of match should have returned by now. If not received and processed, take appropriate action.", 0, date, True, TIKL_note_text)
		script_end_procedure_with_error_report("Match has been acted on. Please take any additional action needed for your case.")
	END IF
'END IF
