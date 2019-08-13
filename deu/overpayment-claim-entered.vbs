'GATHERING STATS===========================================================================================
name_of_script = "ACTIONS - DEU-OVERPAYMENT CLAIM ENTERED.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 500
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
'END FUNCTIONS LIBRARY BLOCK================================================================================================
FUNCTION write_variable_in_CCOL_note_test(variable)
    ''--- This function writes a variable in CCOL note
    '~~~~~ variable: information to be entered into CASE note from script/edit box
    '===== Keywords: MAXIS, CASE note
    If trim(variable) <> "" THEN
    	EMGetCursor noting_row, noting_col						'Needs to get the row and col to start. Doesn't need to get it in the array function because that uses EMWriteScreen.
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
						'msgbox next_page
    				'ELSE
    					Do
    						EMReadScreen character_test, 40, noting_row, 3 	'Reads a single character at the noting row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond case note range).
    						character_test = trim(character_test)
    						If character_test <> "" then noting_row = noting_row + 1
    					Loop until character_test = ""
    				Else
						next_page = FALSE
						'msgbox next_page
						noting_row = 5											'Resets this variable to 4 if we did not need a brand new note.
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
    		'Writes the word and a space using EMWriteScreen
    		EMWriteScreen replace(word, ";", "") & " ", noting_row, noting_col
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
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")

call changelog_update("08/05/2019", "Updated the term claim referral to use the action taken on MISC.", "MiKayla Handley")

CALL changelog_update("04/15/2019", "Updated script to copy case note to CCOL.", "MiKayla Handley, Hennepin County")
CALL changelog_update("01/30/2019", "Updated script to add areas for multiple claims based on request.", "MiKayla Handley, Hennepin County")
CALL changelog_update("04/02/2018", "Updates to fraud referral for the case note.", "MiKayla Handley, Hennepin County")
CALL changelog_update("07/23/2018", "Updated script to correct version and added case note to email for HC matches.", "MiKayla Handley, Hennepin County")
CALL changelog_update("04/02/2018", "Updates to fraud referral for the case note.", "MiKayla Handley, Hennepin County")
CALL changelog_update("02/27/2018", "Added income received date.", "MiKayla Handley, Hennepin County")
CALL changelog_update("01/02/2018", "Corrected IEVS match error due to new year.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/11/2017", "Initial version.", "MiKayla Handley, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)
memb_number = "01"
discovery_date = date & ""

back_to_self
'--------------------------------------------------------------------Dialog
BeginDialog overpayment_dialog, 0, 0, 361, 280, "Overpayment Claim Entered"
  EditBox 60, 5, 40, 15, MAXIS_case_number
  EditBox 140, 5, 20, 15, memb_number
  EditBox 230, 5, 20, 15, OT_resp_memb
  DropListBox 310, 5, 45, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", fraud_referral
  EditBox 60, 25, 40, 15, discovery_date
  DropListBox 210, 25, 40, 15, "Select:"+chr(9)+"WAGE"+chr(9)+"BEER", IEVS_type
  DropListBox 310, 25, 45, 15, "Select:"+chr(9)+"1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"YEAR"+chr(9)+"LAST YEAR"+chr(9)+"OTHER", select_quarter
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
  EditBox 40, 155, 30, 15, HC_from
  EditBox 90, 155, 30, 15, HC_to
  EditBox 160, 155, 50, 15, HC_claim_number
  EditBox 235, 155, 45, 15, HC_claim_amount
  EditBox 100, 175, 20, 15, HC_resp_memb
  EditBox 235, 175, 45, 15, Fed_HC_AMT
  EditBox 70, 200, 160, 15, income_source
  CheckBox 235, 205, 120, 10, "Earned income disregard allowed", EI_checkbox
  EditBox 70, 220, 160, 15, EVF_used
  EditBox 310, 220, 45, 15, income_rcvd_date
  EditBox 70, 240, 285, 15, Reason_OP
  ButtonGroup ButtonPressed
    OkButton 260, 260, 45, 15
    CancelButton 310, 260, 45, 15
  Text 5, 10, 50, 10, "Case number: "
  Text 110, 10, 30, 10, "Memb #:"
  Text 170, 10, 60, 10, "OT resp. Memb #:"
  Text 260, 10, 50, 10, "Fraud referral:"
  Text 5, 30, 55, 10, "Discovery date: "
  Text 170, 30, 40, 10, "Match type:"
  Text 260, 30, 45, 10, "Match period:"
  GroupBox 5, 45, 350, 100, "Overpayment Information"
  Text 15, 70, 30, 10, "Program:"
  Text 105, 70, 20, 10, "From:"
  Text 165, 70, 10, 10, "To:"
  Text 215, 70, 25, 10, "Claim #"
  Text 285, 70, 20, 10, "AMT:"
  Text 130, 55, 30, 10, "(MM/YY)"
  Text 180, 55, 30, 10, "(MM/YY)"
  Text 15, 90, 30, 10, "Program:"
  Text 105, 90, 20, 10, "From:"
  Text 165, 90, 10, 10, "To:"
  Text 215, 90, 25, 10, "Claim #"
  Text 285, 90, 20, 10, "AMT:"
  Text 15, 110, 30, 10, "Program:"
  Text 105, 110, 20, 10, "From:"
  Text 165, 110, 10, 10, "To:"
  Text 215, 110, 25, 10, "Claim #"
  Text 285, 110, 20, 10, "AMT:"
  Text 15, 90, 30, 10, "Program:"
  Text 15, 130, 30, 10, "Program:"
  Text 105, 130, 20, 10, "From:"
  Text 165, 130, 10, 10, "To:"
  Text 215, 130, 25, 10, "Claim #"
  Text 285, 130, 20, 10, "AMT:"
  Text 5, 225, 65, 10, "Income verif used:"
  Text 15, 180, 80, 10, "HC OT resp. Memb(s) #:"
  Text 160, 180, 75, 10, "Total federal HC AMT:"
  Text 30, 245, 40, 10, "OP reason:"
  Text 245, 225, 60, 10, "Date income rcvd: "
  Text 215, 160, 20, 10, "AMT:"
  Text 15, 205, 50, 10, "Income source:"
  Text 15, 160, 20, 10, "From:"
  Text 130, 160, 25, 10, "Claim #"
  Text 75, 160, 10, 10, "To:"
  GroupBox 5, 145, 350, 50, "HC Programs Only"
  Text 15, 70, 30, 10, "Program:"
  Text 165, 70, 10, 10, "To:"
  GroupBox 5, 45, 350, 100, "Overpayment Information"
  Text 105, 70, 20, 10, "From:"
  CheckBox 5, 260, 250, 15, "DHS 2776E Cash (agency) Error Overpayment Worksheet form completed", ECF_checkbox
EndDialog

Do
	err_msg = ""
	dialog overpayment_dialog
	cancel_confirmation
	IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
    IF select_quarter = "Select:" THEN err_msg = err_msg & vbnewline & "* You must select a match period entry."
	IF fraud_referral = "Select:" THEN err_msg = err_msg & vbnewline & "* You must select a fraud referral entry."
	IF trim(Reason_OP) = "" or len(Reason_OP) < 5 THEN err_msg = err_msg & vbnewline & "* You must enter a reason for the overpayment please provide as much detail as possible (min 5)."
    	'IF OP_program = "Select:"THEN err_msg = err_msg & vbNewLine &  "* Please enter the program for the overpayment."
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
    	IF EVF_used = "" then err_msg = err_msg & vbNewLine & "* Please enter verification used for the income recieved. If no verification was received enter N/A."
	'IF isdate(income_rcvd_date) = False or income_rcvd_date = "" then err_msg = err_msg & vbNewLine & "* Please enter a valid date for the income recieved."
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
LOOP UNTIL err_msg = ""
CALL check_for_password_without_transmit(are_we_passworded_out)

'----------------------------------------------------------------------------------------------------STAT
CALL navigate_to_MAXIS_screen("STAT", "MEMB")
EMwritescreen memb_number, 20, 76
EMReadScreen first_name, 12, 6, 63
first_name = replace(first_name, "_", "")
first_name = trim(first_name)
transmit
EMReadscreen SSN_number_read, 11, 7, 42
SSN_number_read = replace(SSN_number_read, " ", "")
CALL navigate_to_MAXIS_screen("INFC" , "____")
CALL write_value_and_transmit("IEVP", 20, 71)
CALL write_value_and_transmit(SSN_number_read, 3, 63) '
EMReadScreen edit_error, 2, 24, 2
edit_error = trim(edit_error)
IF edit_error <> "" THEN script_end_procedure_with_error_report("No IEVS matches and/ or could not access IEVP.")
'---------------------------------------------------------------------------------------------Chosing the match to clear'
Row = 7
DO
	EMReadScreen IEVS_match, 11, row, 47
	IF trim(IEVS_match) = "" THEN script_end_procedure_with_error_report("IEVS match for the selected period could not be found. The script will now end.")
	ievp_info_confirmation = MsgBox("Press YES to confirm this is the match you wish to act on." & vbNewLine & "For the next match, press NO." & vbNewLine & vbNewLine & _
	"   " & IEVS_match, vbYesNoCancel, "Please confirm this match")
	'msgbox IEVS_match
	IF ievp_info_confirmation = vbNo THEN
		row = row + 1
	'msgbox "row: " & row
		IF row = 17 THEN
			PF8
			row = 7
		END IF
	END IF
	IF ievp_info_confirmation = vbCancel THEN script_end_procedure_with_error_report("The script has ended. The match has not been acted on.")
	IF ievp_info_confirmation = vbYes THEN 	EXIT DO
LOOP UNTIL ievp_info_confirmation = vbYes

EMReadScreen multiple_match, 11, row + 1, 47
IF multiple_match = IEVS_match THEN
	msgbox("More than one match exists for this time period. Determine the match you'd like to clear, and put your cursor in front of that match." & vbcr & "Press OK once match is determined.")
	EMSendKey "U"
	transmit
ELSE
	CALL write_value_and_transmit("U", row, 3)   'navigates to IULA
END IF

'----------------------------------------------------------------------------------------------------IULA
'Entering the IEVS match & reading the difference notice to ensure this has been sent
'Reading potential errors for out-of-county cases
EMReadScreen OutOfCounty_error, 12, 24, 2
IF OutOfCounty_error = "MATCH IS NOT" then
	script_end_procedure_with_error_report("Out-of-county case. Cannot update.")
Else
	IF IEVS_type = "WAGE" then
		EMReadScreen quarter, 1, 8, 14
		EMReadScreen IEVS_year, 4, 8, 22
		If quarter <> select_quarter then script_end_procedure_with_error_report("Match period does not match the selected match period. The script will now end.")
	Elseif IEVS_type <> "WAGE" THEN
		EMReadScreen IEVS_year, 4, 8, 15
	End if
End if

'----------------------------------------------------------------------------------------------------Client name
EMReadScreen client_name, 35, 5, 24
'Formatting the client name for the spreadsheet
client_name = trim(client_name)                         'trimming the client name
if instr(client_name, ",") then    						'Most cases have both last name and 1st name. This separates the two names
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
'----------------------------------------------------------------------------------------------------Employer info & diff notice info
EMReadScreen income_source, 74, 8, 37
income_source = trim(income_source)
length = len(income_source)		'establishing the length of the variable
IF instr(income_source, " AMOUNT: $") THEN
    position = InStr(income_source, " AMOUNT: $")    		      'sets the position at the deliminator
    income_source = Left(income_source, position)  'establishes employer as being before the deliminator
Elseif instr(income_source, " AMT:") THEN 					  'establishing the length of the variable
    position = InStr(income_source, " AMT: $")    		      'sets the position at the deliminator
    income_source = Left(income_source, position)  'establishes employer as being before the deliminator
Else
    income_source = income_source	'catch all variable
END IF
EMReadScreen diff_notice, 1, 14, 37
EMReadScreen diff_date, 10, 14, 68
diff_date = trim(diff_date)
If diff_date <> "" then diff_date = replace(diff_date, " ", "/")
IF IEVS_type = "UNVI" THEN income_source = replace(income_source, "")
PF3		'exiting IULA, helps prevent errors when going to the case note
'-----------------------------------------------------------------------------------'for the case notes
programs = ""
IF instr(Active_Programs, "D") then programs = programs & "DWP, "
IF instr(Active_Programs, "F") then programs = programs & "Food Support, "
IF instr(Active_Programs, "H") then programs = programs & "Health Care, "
IF instr(Active_Programs, "M") then programs = programs & "Medical Assistance, "
IF instr(Active_Programs, "S") then programs = programs & "MFIP, "
IF other_programs = CHECKED THEN programs = "Food Support, "
'trims excess spaces of programs
programs = trim(programs)
'takes the last comma off of programs when auto filled into dialog
IF right(programs, 1) = "," THEN programs = left(programs, len(programs) - 1)
If IEVS_type = "WAGE" THEN
	'Updated IEVS_period to write into case note
	If select_quarter = "1" then IEVS_quarter = "1ST"
	If select_quarter = "2" then IEVS_quarter = "2ND"
	If select_quarter = "3" then IEVS_quarter = "3RD"
	If select_quarter = "4" then IEVS_quarter = "4TH"
End if

IF IEVS_type = "UNVI" THEN type_match = "U"
IF IEVS_type = "BEER" THEN type_match = "B"

IEVS_period = replace(IEVS_period, "/", " to ")
Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days requested for HEADER of casenote'
PF3 'back to the DAIL'
IF OP_program = "FS" or OP_program_II = "FS" or OP_program_III = "FS" or OP_program_IV = "FS" or OP_program = "MF" or OP_program_II = "MF" or OP_program_III = "MF" or OP_program_IV = "MF" THEN
	'Going to the MISC panel to add claim referral tracking information
	Call navigate_to_MAXIS_screen ("STAT", "MISC")
	Row = 6
	EmReadScreen panel_number, 1, 02, 73
	If panel_number = "0" then
		EMWriteScreen "NN", 20,79
		TRANSMIT
		'CHECKING FOR MAXIS PROGRAMS ARE INACTIVE'
		EmReadScreen MISC_error_check,  74, 24, 02
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

	'writing in the action taken and date to the MISC panel
	PF9
	EMWriteScreen "Determination-OP Entered", Row, 30
	EMWriteScreen date, Row, 66
	TRANSMIT
	start_a_blank_case_note
	Call write_variable_in_case_note("-----Claim Referral Tracking - Claim Determination-----")
	IF case_note_only = TRUE THEN Call write_variable_in_case_note("Maxis case is inactive unable to add or update MISC panel")
	Call write_bullet_and_variable_in_case_note("Action Date", date)
	Call write_bullet_and_variable_in_case_note("Program(s)", programs)
	Call write_variable_in_case_note("* Entries for these potential claims must be retained until further notice.")
	Call write_variable_in_case_note("-----")
	Call write_variable_in_case_note(worker_signature)
	PF3
END IF
'-----------------------------------------------------------------------------------------CASENOTE
start_a_blank_CASE_NOTE
IF IEVS_type = "WAGE" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH"  & "(" & first_name & ") CLEARED CC-CLAIM ENTERED-----")
IF IEVS_type = "BEER" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_year & " NON-WAGE MATCH(" & match_type & ") " & "(" & first_name & ") CLEARED CC-CLAIM ENTERED-----")
IF IEVS_type = "UBEN" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_month & " NON-WAGE MATCH(" & match_type & ") " & "(" & first_name & ") CLEARED CC-CLAIM ENTERED-----")
IF IEVS_type = "UNVI" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_year & " NON-WAGE MATCH(" & match_type & ") " & "(" & first_name & ") CLEARED CC-CLAIM ENTERED-----")
CALL write_bullet_and_variable_in_CASE_NOTE("Discovery date", discovery_date)
'CALL write_bullet_and_variable_in_CASE_NOTE("Period", IEVS_period)
CALL write_bullet_and_variable_in_CASE_NOTE("Active Programs", programs)
CALL write_bullet_and_variable_in_CASE_NOTE("Source of income", income_source)
Call write_variable_in_CASE_NOTE("----- ----- ----- ----- -----")
Call write_variable_in_CASE_NOTE(OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number & " Amt $" & Claim_amount)
IF OP_program_II <> "Select:" then Call write_variable_in_CASE_NOTE(OP_program_II & " Overpayment " & OP_from_II & " through " & OP_to_II & " Claim # " & Claim_number_II & " Amt $" & Claim_amount_II)
IF OP_program_III <> "Select:" then	Call write_variable_in_CASE_NOTE(OP_program_III & " Overpayment " & OP_from_III & " through " & OP_to_III & " Claim # " & Claim_number_III & " Amt $" & Claim_amount_III)
IF OP_program_IV <> "Select:" then Call write_variable_in_CASE_NOTE(OP_program_IV & " Overpayment " & OP_from_IV & " through " & OP_to_IV & " Claim # " & Claim_number_IV & " Amt $" & Claim_amount_IV)
IF HC_claim_number <> "" THEN
	Call write_variable_in_case_note("HC OVERPAYMENT " & HC_from & " through " & HC_to & " Claim #" & HC_claim_number & " Amt $" & HC_Claim_amount)
	Call write_bullet_and_variable_in_CASE_NOTE("Health Care responsible members", HC_resp_memb)
	Call write_bullet_and_variable_in_CASE_NOTE("Total Federal Health Care amount", Fed_HC_AMT)
	Call write_variable_in_CASE_NOTE("* Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
END IF
IF EI_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Earned Income Disregard Allowed")
IF EI_checkbox = UNCHECKED THEN CALL write_variable_in_case_note("* Earned Income Disregard Not Allowed")
CALL write_bullet_and_variable_in_case_note("Fraud referral made", fraud_referral)
CALL write_bullet_and_variable_in_case_note("Income verification received", EVF_used)
CALL write_bullet_and_variable_in_case_note("Date verification received", income_rcvd_date)
CALL write_bullet_and_variable_in_case_note("Reason for overpayment", Reason_OP)
CALL write_bullet_and_variable_in_case_note("Other responsible member(s)", OT_resp_memb)
IF ECF_checkbox = CHECKED THEN CALL write_variable_in_CASE_NOTE("* DHS 2776E – Agency Cash Error Overpayment Worksheet form completed in ECF")
CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- ----- ----- -----")
CALL write_variable_in_CASE_NOTE("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1")
PF3 'to save casenote'

IF HC_claim_number <> "" THEN
	EmWriteScreen "x", 5, 3
	Transmit
	note_row = 4			'Beginning of the case notes
	Do 						'Read each line
		EMReadScreen note_line, 76, note_row, 3
		note_line = trim(note_line)
		If trim(note_line) = "" Then Exit Do		'Any blank line indicates the end of the case note because there can be no blank lines in a note
		message_array = message_array & note_line & vbcr		'putting the lines together
		note_row = note_row + 1
		If note_row = 18 then 									'End of a single page of the case note
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

'---------------------------------------------------------------writing the CCOL case note'
msgbox "Navigating to CCOL to add case note, please contact the BlueZone Scripts team with any concerns."
Call navigate_to_MAXIS_screen("CCOL", "CLSM")
EMWriteScreen Claim_number, 4, 9
TRANSMIT
'NO CLAIMS WERE FOUND FOR THIS CASE, PROGRAM, AND STATUS
EMReadScreen error_check, 75, 24, 2	'making sure we can actually update this case.
error_check = trim(error_check)
If error_check <> "" then script_end_procedure_with_error_report(error_check & ". Unable to update this case. Please review case, and run the script again if applicable.")
PF4
EMReadScreen existing_case_note, 1, 5, 6
IF existing_case_note = "" THEN
	PF4
ELSE
	PF9
END IF
'-----------------------------------------------------------------------------------------CCOL CASENOTE
IF IEVS_type = "WAGE" THEN CALL write_variable_in_CCOL_note_test("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH"  & "(" & first_name & ") CLEARED CC-CLAIM ENTERED-----")
IF IEVS_type = "BEER" THEN CALL write_variable_in_CCOL_note_test("-----" & IEVS_year & " NON-WAGE MATCH(" & match_type & ") " & "(" & first_name & ") CLEARED CC-CLAIM ENTERED-----")
IF IEVS_type = "UBEN" THEN CALL write_variable_in_CCOL_note_test("-----" & IEVS_month & " NON-WAGE MATCH(" & match_type & ") " & "(" & first_name & ") CLEARED CC-CLAIM ENTERED-----")
IF IEVS_type = "UNVI" THEN CALL write_variable_in_CCOL_note_test("-----" & IEVS_year & " NON-WAGE MATCH(" & match_type & ") " & "(" & first_name & ") CLEARED CC-CLAIM ENTERED-----")
CALL write_bullet_and_variable_in_CCOL_note_test("Discovery date", discovery_date)
CALL write_bullet_and_variable_in_CCOL_note_test("Active Programs", programs)
CALL write_bullet_and_variable_in_CCOL_note_test("Source of income", income_source)
Call write_variable_in_CCOL_note_test("----- ----- ----- ----- -----")
IF OP_program <> "Select:" then Call write_variable_in_CCOL_note_test(OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number & " Amt $" & Claim_amount)
IF OP_program_II <> "Select:" then Call write_variable_in_CCOL_note_test(OP_program_II & " Overpayment " & OP_from_II & " through " & OP_to_II & " Claim # " & Claim_number_II & " Amt $" & Claim_amount_II)
IF OP_program_III <> "Select:" then Call write_variable_in_CCOL_note_test(OP_program_III & " Overpayment " & OP_from_III & " through " & OP_to_III & " Claim # " & Claim_number_III & " Amt $" & Claim_amount_III)
IF OP_program_IV <> "Select:" then Call write_variable_in_CCOL_note_test(OP_program_IV & " Overpayment " & OP_from_IV & " through " & OP_to_IV & " Claim # " & Claim_number_IV & " Amt $" & Claim_amount_IV)
IF HC_claim_number <> "" THEN
Call write_variable_in_CCOL_note_test("HC OVERPAYMENT " & HC_from & " through " & HC_to & " Claim #" & HC_claim_number & " Amt $" & HC_Claim_amount)
	Call write_bullet_and_variable_in_CCOL_note_test("Health Care responsible members", HC_resp_memb)
	Call write_bullet_and_variable_in_CCOL_note_test("Total Federal Health Care amount", Fed_HC_AMT)
	Call write_variable_in_CCOL_note_test("* Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
END IF
IF EI_checkbox = CHECKED THEN CALL write_variable_in_CCOL_note_test("* Earned Income Disregard Allowed")
IF EI_checkbox = UNCHECKED THEN CALL write_variable_in_CCOL_note_test("* Earned Income Disregard Not Allowed")
CALL write_bullet_and_variable_in_CCOL_note_test("Fraud referral made", fraud_referral)
CALL write_bullet_and_variable_in_CCOL_note_test("Income verification received", EVF_used)
CALL write_bullet_and_variable_in_CCOL_note_test("Date verification received", income_rcvd_date)
CALL write_bullet_and_variable_in_CCOL_note_test("Reason for overpayment", Reason_OP)
CALL write_bullet_and_variable_in_CCOL_note_test("Other responsible member(s)", OT_resp_memb)
IF ECF_checkbox = CHECKED THEN CALL write_variable_in_CCOL_note_test("* DHS 2776E - Agency Cash Error Overpayment Worksheet form completed in ECF")
CALL write_variable_in_CCOL_note_test("----- ----- ----- ----- ----- ----- -----")
CALL write_variable_in_CCOL_note_test("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1")
PF3'

script_end_procedure_with_error_report("Overpayment case note entered and copied to CCOL please review case note to ensure accuracy.")
