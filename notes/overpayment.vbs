'GATHERING STATS===========================================================================================
name_of_script = "NOTES - OVERPAYMENT CLAIM ENTERED.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 500
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
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("05/18/2020", "GitHub issue #381 Added Requested Claim Adjustment per project request.", "MiKayla Handley")
call changelog_update("03/04/2020", "Removed agency error OP worksheet as the form is now obsolete.", "MiKayla Handley")
call changelog_update("08/05/2019", "Updated the term claim referral to use the action taken on MISC as well as to read for active programs.", "MiKayla Handley")
CALL changelog_update("04/15/2019", "Updated script to copy case note to CCOL.", "MiKayla Handley, Hennepin County")
CALL changelog_update("07/23/2018", "Updated script to correct version and added case note to email for HC matches.", "MiKayla Handley, Hennepin County")
CALL changelog_update("04/02/2018", "Updates to fraud referral for the case note.", "MiKayla Handley, Hennepin County")
CALL changelog_update("02/27/2018", "Added income received date.", "MiKayla Handley, Hennepin County")
CALL changelog_update("01/02/2018", "Corrected IEVS match error due to new year.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/11/2017", "Initial version.", "MiKayla Handley, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'---------------------------------------------------------------------------------------The script
'Grabs the case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)
memb_number = "01" 'defaults to 01'
discovery_date = "" & date
'-------------------------------------------------------------------------------------------------DIALOG
BeginDialog Dialog1, 0, 0, 161, 85, "Overpayment/Claim"
  EditBox 55, 5, 45, 15, MAXIS_case_number
  DropListBox 55, 25, 105, 15, "Select One:"+chr(9)+"Intial Overpayment/Claim"+chr(9)+"Requested Claim Adjustment", claim_actions
  EditBox 55, 45, 105, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 65, 65, 45, 15
    CancelButton 115, 65, 45, 15
  Text 5, 30, 50, 10, "Claim Action:"
  Text 5, 10, 50, 10, "Case Number:"
  Text 5, 50, 40, 10, "Worker Sig:"
EndDialog

'Runs the first dialog - which confirms the case number
Do
	Do
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
      	IF IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
        IF claim_actions = "Select One:" then err_msg = err_msg & vbNewLine & "* Please select what type of appeal action the client is claiming."
		IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "* Please enter your worker signature, for help see utilities. "
		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = ""
CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

IF claim_actions = "Intial Overpayment/Claim" THEN
    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 361, 280, "Overpayment Claim Entered"
      EditBox 60, 5, 40, 15, discovery_date
      EditBox 140, 5, 20, 15, memb_number
      EditBox 235, 5, 20, 15, OT_resp_memb
      DropListBox 310, 5, 45, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", fraud_referral
      DropListBox 50, 40, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program
      EditBox 130, 40, 30, 15, OP_from
      EditBox 180, 40, 30, 15, OP_to
      EditBox 245, 40, 35, 15, Claim_number
      EditBox 305, 40, 45, 15, Claim_amount
      DropListBox 50, 60, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program_II
      EditBox 130, 60, 30, 15, OP_from_II
      EditBox 180, 60, 30, 15, OP_to_II
      EditBox 245, 60, 35, 15, Claim_number_II
      EditBox 305, 60, 45, 15, Claim_amount_II
      DropListBox 50, 80, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program_III
      EditBox 130, 80, 30, 15, OP_from_III
      EditBox 180, 80, 30, 15, OP_to_III
      EditBox 245, 80, 35, 15, claim_number_III
      EditBox 305, 80, 45, 15, Claim_amount_III
      DropListBox 50, 100, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program_IV
      EditBox 130, 100, 30, 15, OP_from_IV
      EditBox 180, 100, 30, 15, OP_to_IV
      EditBox 245, 100, 35, 15, claim_number_IV
      EditBox 305, 100, 45, 15, Claim_amount_IV
      EditBox 40, 135, 30, 15, HC_from
      EditBox 90, 135, 30, 15, HC_to
      EditBox 160, 135, 50, 15, HC_claim_number
      EditBox 235, 135, 45, 15, HC_claim_amount
      EditBox 40, 155, 30, 15, HC_from_II
      EditBox 90, 155, 30, 15, HC_to_II
      EditBox 160, 155, 50, 15, HC_claim_number_II
      EditBox 235, 155, 45, 15, HC_claim_amount_II
      EditBox 100, 175, 20, 15, HC_resp_memb
      EditBox 235, 175, 45, 15, Fed_HC_AMT
      EditBox 70, 200, 160, 15, income_source
      CheckBox 235, 205, 120, 10, "Earned income disregard allowed", EI_checkbox
      EditBox 70, 220, 160, 15, EVF_used
      EditBox 310, 220, 45, 15, income_rcvd_date
      EditBox 70, 240, 285, 15, Reason_OP
      'CheckBox 5, 265, 240, 10, "DHS 2776E Cash (agency) Error OP Worksheet form completed in ECF", ECF_checkbox
       Text 5, 10, 55, 10, "Discovery date:"
       Text 110, 10, 30, 10, "Memb #:"
       Text 165, 10, 70, 10, "Other resp. memb #:"
       Text 260, 10, 50, 10, "Fraud referral:"
       GroupBox 5, 25, 350, 95, "Overpayment Information"
       Text 130, 30, 30, 10, "(MM/YY)"
       Text 180, 30, 30, 10, "(MM/YY)"
       Text 15, 45, 30, 10, "Program:"
       Text 105, 45, 20, 10, "From:"
       Text 165, 45, 10, 10, "To:"
       Text 215, 45, 25, 10, "Claim #"
       Text 285, 45, 20, 10, "AMT:"
       Text 15, 65, 30, 10, "Program:"
       Text 105, 65, 20, 10, "From:"
       Text 165, 65, 10, 10, "To:"
       Text 215, 65, 25, 10, "Claim #"
       Text 285, 65, 20, 10, "AMT:"
       Text 15, 85, 30, 10, "Program:"
       Text 105, 85, 20, 10, "From:"
       Text 165, 85, 10, 10, "To:"
       Text 215, 85, 25, 10, "Claim #"
       Text 285, 85, 20, 10, "AMT:"
       Text 15, 105, 30, 10, "Program:"
       Text 105, 105, 20, 10, "From:"
       Text 165, 105, 10, 10, "To:"
       Text 215, 105, 25, 10, "Claim #"
       Text 285, 105, 20, 10, "AMT:"
       ButtonGroup ButtonPressed
         OkButton 260, 260, 45, 15
         CancelButton 310, 260, 45, 15
       GroupBox 5, 125, 350, 70, "HC Programs Only"
       Text 15, 160, 20, 10, "From:"
       Text 75, 160, 10, 10, "To:"
       Text 130, 160, 25, 10, "Claim #"
       Text 215, 160, 20, 10, "AMT:"
       Text 15, 140, 20, 10, "From:"
       Text 75, 140, 10, 10, "To:"
       Text 130, 140, 25, 10, "Claim #"
       Text 215, 140, 20, 10, "AMT:"
       Text 15, 180, 80, 10, "HC OT resp. Memb(s) #:"
       Text 160, 180, 75, 10, "Total federal HC AMT:"
       Text 5, 205, 50, 10, "Income source:"
       Text 5, 225, 65, 10, "Income verif used:"
       Text 5, 245, 60, 10, "Reason for claim:"
       Text 235, 225, 75, 10, "Date income received:"
     EndDialog

    Do
        Do
        	err_msg = ""
        	dialog Dialog1
        	cancel_confirmation
        	IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
        	IF fraud_referral = "Select:" THEN err_msg = err_msg & vbnewline & "* You must select a fraud referral entry."
        	IF trim(Reason_OP) = "" or len(Reason_OP) < 5 THEN err_msg = err_msg & vbnewline & "* You must enter a reason for the overpayment please provide as much detail as possible (min 5)."
        	'IF OP_program = "Select:"THEN err_msg = err_msg & vbNewLine &  "* Please enter the program for the overpayment."
        	IF OP_program_II <> "Select:" THEN
				IF OP_from_II = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the period the overpayment occurred, from month and year (MM/YY)."
				IF OP_to_II = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the period the overpayment occurred, to month and year (MM/YY)."
    	   		IF Claim_number_II = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the claim number."
        		IF Claim_amount_II = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the amount of claim."
        	END IF
    	    IF OP_program_III <> "Select:" THEN
				IF OP_from_III = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the period the overpayment occurred, from month and year (MM/YY)."
				IF OP_to_III = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the period the overpayment occurred, to month and year (MM/YY)."
    	    	IF Claim_number_III = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the claim number."
    	    	IF Claim_amount_III = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the amount of claim."
    	    END IF
    	    IF OP_program_IV <> "Select:" THEN
				IF OP_from_IV = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the period the overpayment occurred, from month and year (MM/YY)."
				IF OP_to_IV = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the period the overpayment occurred, to month and year (MM/YY)."
    	    	IF Claim_number_IV = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the claim number."
    	    	IF Claim_amount_IV = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the amount of claim."
    	    END IF
            	IF HC_claim_number <> "" THEN
            		IF HC_from = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the month and year overpayment started."
            		IF HC_to = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the month and year overpayment ended."
            		IF HC_claim_amount = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the amount of claim."
            	END IF
            	IF EVF_used = "" then err_msg = err_msg & vbNewLine & "* Please enter verification used for the income received. If no verification was received enter N/A."
            	'IF isdate(income_rcvd_date) = False or income_rcvd_date = "" then err_msg = err_msg & vbNewLine & "* Please enter a valid date for the income received."
            	IF ECF_checkbox = UNCHECKED and OP_program = "MF" THEN err_msg = err_msg & vbNewLine &  "* Please ensure you are entering the FS OP Determination form in ECF and check the appropriate box."
            	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
        LOOP UNTIL err_msg = ""
        CALL check_for_password(are_we_passworded_out)
    Loop until are_we_passworded_out = False
	'---------------------------------------------------------------------------------------------'client information
	CALL navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_priv)
	IF is_this_priv = TRUE THEN script_end_procedure("This case is privileged, the script will now end.")

    EMwritescreen MEMB_number, 20, 76
	TRANSMIT
	EMReadscreen panel_MEMB_number, 2, 4, 33
	'MsgBox panel_MEMB_number & " ~ " &  MEMB_number
	IF panel_MEMB_number <> MEMB_number THEN script_end_procedure_with_error_report("This MEMB was not found, the script will now end.")
	EMReadscreen last_name, 25, 6, 30
	EMReadscreen first_name, 12, 6, 63
	EMReadscreen mid_initial, 1, 6, 79

	last_name = trim(replace(last_name, "_", ""))
	first_name = trim(replace(first_name, "_", ""))
	mid_initial = replace(mid_initial, "_", "")
	client_name = MEMB_number & " - " & last_name &  ", " & first_name & " " & mid_initial
	'MsgBox client_name
    client_name = trim(client_name)
    '-----------------------------------------------------------------------------------------CASENOTE
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
    	IF row = 17 then MsgBox("There is not a blank field in the MISC panel. Please delete a line(s), and run script again or update manually.")

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
    Call write_variable_in_CASE_NOTE("OVERPAYMENT CLAIM ENTERED" & " (" & client_name & ") ")
    CALL write_bullet_and_variable_in_CASE_NOTE("Discovery date", discovery_date)
    CALL write_bullet_and_variable_in_CASE_NOTE("Active Programs", programs)
    CALL write_bullet_and_variable_in_CASE_NOTE("Source of income", income_source)
    Call write_variable_in_CASE_NOTE("----- ----- ----- ----- -----")
    Call write_variable_in_CASE_NOTE(OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number & " Amt $" & Claim_amount)
    IF OP_program_II <> "Select:" then Call write_variable_in_CASE_NOTE(OP_program_II & " Overpayment " & OP_from_II & " through " & OP_to_II & " Claim # " & Claim_number_II & " Amt $" & Claim_amount_II)
    IF OP_program_III <> "Select:" then	Call write_variable_in_CASE_NOTE(OP_program_III & " Overpayment " & OP_from_III & " through " & OP_to_III & " Claim # " & Claim_number_III & " Amt $" & Claim_amount_III)
    IF OP_program_IV <> "Select:" then Call write_variable_in_CASE_NOTE(OP_program_IV & " Overpayment " & OP_from_IV & " through " & OP_to_IV & " Claim # " & Claim_number_IV & " Amt $" & Claim_amount_IV)
    IF HC_claim_number <> "" THEN
    	Call write_variable_in_case_note("HC OVERPAYMENT " & HC_from & " through " & HC_to & " Claim #" & HC_claim_number & " Amount $" & HC_Claim_amount)
    	Call write_bullet_and_variable_in_CASE_NOTE("Health Care responsible members", HC_resp_memb)
    	Call write_bullet_and_variable_in_CASE_NOTE("Total Federal Health Care amount", Fed_HC_AMT)
    	Call write_variable_in_CASE_NOTE("* Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
    END IF
    IF HC_claim_number_II <> "" THEN
    	Call write_variable_in_case_note("HC OVERPAYMENT " & HC_from_II & " through " & HC_to_II & " Claim #" & HC_claim_number_II & " Amount $" & HC_claim_amount_II)
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
    'IF ECF_checkbox = CHECKED THEN CALL write_variable_in_CASE_NOTE("* DHS 2776E – Agency Cash Error Overpayment Worksheet form completed in ECF")
    CALL write_variable_in_CASE_NOTE("----- ----- -----")
    CALL write_variable_in_CASE_NOTE(worker_signature)
    PF3 'to save casenote'

    IF HC_claim_number <> "" THEN
    	EmWriteScreen "x", 5, 3
    	TRANSMIT
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

    Call write_variable_in_CCOL_note_test("OVERPAYMENT CLAIM ENTERED" & " (" & client_name & ") ")
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
    IF HC_claim_number_II <> "" THEN
    	Call write_variable_in_CCOL_note_test("HC OVERPAYMENT " & HC_from_II & " through " & HC_to_II & " Claim #" & HC_claim_number_II & " Amount $" & HC_claim_amount_II)
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
    CALL write_variable_in_CCOL_note_test("----- ----- -----")
    CALL write_variable_in_CCOL_note_test(worker_signature)
    PF3 'to save casenote'
END IF

IF claim_actions = "Requested Claim Adjustment" THEN
    BeginDialog Dialog1, 0, 0, 226, 125, "Requested Claim Adjustment"
      EditBox 60, 5, 50, 15, claim_number
      EditBox 180, 5, 40, 15, original_claim_amount
      EditBox 180, 25, 40, 15, corrected_claim_amount
      EditBox 60, 45, 35, 15, OP_from
      EditBox 115, 45, 35, 15, OP_to
      EditBox 80, 65, 140, 15, reason_correction
      EditBox 60, 85, 160, 15, other_notes
      ButtonGroup ButtonPressed
        OkButton 135, 105, 40, 15
        CancelButton 180, 105, 40, 15
      Text 5, 10, 50, 10, "Claim number:"
      Text 120, 10, 55, 10, "Original Amount:"
      Text 120, 30, 55, 10, "Correct Amount:"
      Text 5, 90, 45, 10, "Other notes:"
      Text 5, 50, 50, 10, "Period    From:"
      Text 100, 50, 15, 10, "To:"
      Text 5, 70, 75, 10, "Reason for correction:"
      Text 60, 35, 35, 10, "(MM/YY)"
    EndDialog

	Do
		Do
			err_msg = ""
			Dialog Dialog1
			cancel_without_confirmation
	      	IF IsNumeric(claim_number) = false THEN err_msg = err_msg & vbNewLine & "* Please enter a valid claim number."
			IF IsNumeric(original_claim_amount) = false THEN err_msg = err_msg & vbNewLine & "* Please enter a valid original claim amount(do not include $)."
	        IF IsNumeric(corrected_claim_amount) = false THEN err_msg = err_msg & vbNewLine & "* Please enter a valid original claim amount(do not include $)."
			IF OP_from = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the period the overpayment occurred, from month and year (MM/YY)."
			IF OP_to = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the period the overpayment occurred, to month and year (MM/YY)."
			IF reason_correction = "" THEN err_msg = err_msg & vbNewLine & "* Please enter the reason the correction is needed. "
			IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
		Loop until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	LOOP UNTIL are_we_passworded_out = false

	'-----------------------------------------------------------------------------------------CASENOTE
    start_a_blank_CASE_NOTE
    Call write_variable_in_CASE_NOTE("Requested Claim Adjustment")
    CALL write_bullet_and_variable_in_CASE_NOTE("Discovery date", discovery_date)
    CALL write_variable_in_CASE_NOTE("* Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number)
	CALL write_bullet_and_variable_in_CASE_NOTE("Original Amount", original_claim_amount)
	Call write_bullet_and_variable_in_CASE_NOTE("Correct Amount",  corrected_claim_amount)
	Call write_bullet_and_variable_in_CASE_NOTE("Reason for correction", reason_correction)
	Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
	CALL write_variable_in_CASE_NOTE("----- ----- -----")
	CALL write_variable_in_CASE_NOTE(worker_signature)
	PF3

	'---------------------------------------------------------------writing the CCOL case note'
    msgbox "Navigating to CCOL to add case note, please contact the BlueZone Scripts Team with any concerns."
    Call navigate_to_MAXIS_screen("CCOL", "CLSM")
    EMWriteScreen Claim_number, 4, 9
    TRANSMIT
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

	Call write_variable_in_CCOL_note_test("Requested Claim Adjustment")
	CALL write_bullet_and_variable_in_CCOL_note_test("Discovery date", discovery_date)
	CALL write_variable_in_CCOL_note_test("* Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number)
	CALL write_bullet_and_variable_in_CCOL_note_test("Original Amount", original_claim_amount)
	Call write_bullet_and_variable_in_CCOL_note_test("Correct Amount", corrected_claim_amount)
	Call write_bullet_and_variable_in_CCOL_note_test("Reason for correction", reason_correction)
	Call write_bullet_and_variable_in_CCOL_note_test("Other notes", other_notes)
	CALL write_variable_in_CCOL_note_test("----- ----- -----")
	CALL write_variable_in_CCOL_note_test(worker_signature)
	PF3
END IF
script_end_procedure_with_error_report("Overpayment case note entered and copied to CCOL please review case note to ensure accuracy.")
