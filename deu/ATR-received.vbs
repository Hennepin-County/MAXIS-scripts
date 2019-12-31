'GATHERING STATS===========================================================================================
name_of_script = "NOTES - DEU-ATR RECEIVED.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 180
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

'================================================================================================CHANGELOG BLOCK
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: CALL changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
CALL changelog_update("01/02/2018", "Corrected IEVS match error due to new year.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/27/2017", "Corrected spelling error.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/11/2017", "Updated to correct case note.", "MiKayla Handley, Hennepin County")
CALL changelog_update("11/07/2017", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'=================================================================================================END CHANGELOG BLOCK
'---------------------------------------------------------------------THE SCRIPT
EMConnect ""
CALL MAXIS_case_number_finder (MAXIS_case_number)
memb_number = "01"
date_received = date & ""
'----------------------------------------------------------------------------------------------------DAIL
EMReadscreen dail_check, 4, 2, 48 'changed from DAIL to view to ensure we are in DAIL/DAIL'
IF dail_check = "DAIL" THEN
	EMSendKey "t"
    EMReadScreen match_type, 4, 6, 6 'read the DAIL msg'
	'msgbox match_type
    IF match_type = "WAGE" or match_type = "BEER" or match_type = "UBEN" or match_type = "UNVI" THEN
    	match_found = TRUE
    ELSE
		match_found = FALSE
		'script_end_procedure("This is not an supported match currently. Please select a WAGE match DAIL, and run the script again.")
    END IF
	IF match_found = TRUE THEN
    	EMReadScreen MAXIS_case_number, 8, 5, 73
		MAXIS_case_number= TRIM(MAXIS_case_number)
		EMReadscreen SSN_number_read, 9, 6, 20

		 '----------------------------------------------------------------------------------------------------IEVP
		'Navigating deeper into the match interface
		CALL write_value_and_transmit("I", 6, 3)   		'navigates to INFC
		CALL write_value_and_transmit("IEVP", 20, 71)   'navigates to IEVP
		TRANSMIT
	    'EMReadScreen err_msg, 7, 24, 2
	    'IF err_msg = "NO IEVS" THEN script_end_procedure_with_error_report("An error occurred in IEVP, please process manually.")'checking for error msg'
	END IF
END IF

IF dail_check <> "DAIL" or match_found = FALSE THEN
    CALL MAXIS_case_number_finder (MAXIS_case_number)
    MEMB_number = "01"
	'-------------------------------------------------------------------------------------------------DIALOG
	Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 131, 65, "Case Number to clear match"
      EditBox 60, 5, 65, 15, MAXIS_case_number
      EditBox 60, 25, 30, 15, MEMB_number
      ButtonGroup ButtonPressed
        OkButton 20, 45, 50, 15
        CancelButton 75, 45, 50, 15
      Text 5, 30, 55, 10, "MEMB Number:"
      Text 5, 10, 50, 10, "Case Number:"
    EndDialog
    DO
    	DO
    		err_msg = ""
    		Dialog Dialog1
    		cancel_without_confirmation
      		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
      		If IsNumeric(MEMB_number) = False or len(MEMB_number) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a valid 2 digit member number."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    	LOOP UNTIL err_msg = ""
    	CALL check_for_password(are_we_passworded_out)
    LOOP UNTIL are_we_passworded_out = false
    CALL navigate_to_MAXIS_screen("STAT", "MEMB")
    EMwritescreen MEMB_number, 20, 76
    TRANSMIT
    EMReadscreen SSN_number_read, 11, 7, 42
    SSN_number_read = replace(SSN_number_read, " ", "")
	CALL navigate_to_MAXIS_screen("INFC" , "____")
	CALL write_value_and_transmit("IEVP", 20, 71)
	CALL write_value_and_transmit(SSN_number_read, 3, 63)
	EmReadscreen err_msg, 50, 24, 02
	err_msg = trim(err_msg)
	'NO IEVS MATCHES FOUND FOR SSN'
	If err_msg <> "" THEN script_end_procedure_with_error_report("*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine)
END IF
'----------------------------------------------------------------------------------------------------selecting the correct wage match
Row = 7
DO
	EMReadScreen IEVS_period, 11, row, 47
	EmReadScreen number_IEVS_type, 3, row, 41
	IF IEVS_period = "" THEN script_end_procedure_with_error_report("A match for the selected period could not be found. The script will now end.")
	ievp_info_confirmation = MsgBox("Press YES to confirm this is the match you wish to act on." & vbNewLine & "For the next match, press NO." & vbNewLine & vbNewLine & _
	" " & "Period: " & IEVS_period & " Type: " & number_IEVS_type, vbYesNoCancel, "Please confirm this match")
	'msgbox IEVS_period
	IF ievp_info_confirmation = vbNo THEN
		row = row + 1
	'msgbox "row: " & row
		IF row = 17 THEN
			PF8
			row = 7
			EMReadScreen IEVS_period, 11, row, 47
		END IF
	END IF
	IF ievp_info_confirmation = vbCancel THEN script_end_procedure_with_error_report ("The script has ended. The match has not been acted on.")
	IF ievp_info_confirmation = vbYes THEN 	EXIT DO
LOOP UNTIL ievp_info_confirmation = vbYes

''---------------------------------------------------------------------Reading potential errors for out-of-county cases
CALL write_value_and_transmit("U", row, 3)   'navigates to IULA
EMReadScreen OutOfCounty_error, 12, 24, 2
IF OutOfCounty_error = "MATCH IS NOT" then
	script_end_procedure_with_error_report("Out-of-county case. Cannot update.")
ELSE
    EMReadScreen number_IEVS_type, 3, 7, 12 'read the DAIL msg'
    IF number_IEVS_type = "A30" THEN match_type = "BNDX"
    'IF number_IEVS_type = "A40" THEN match_type = "SDXS/I"
    IF number_IEVS_type = "A70" THEN match_type = "BEER"
    IF number_IEVS_type = "A80" THEN match_type = "UNVI"
    IF number_IEVS_type = "A60" THEN match_type = "UBEN"
    IF number_IEVS_type = "A50" or number_IEVS_type = "A51"  THEN match_type = "WAGE"

	IF match_type = "WAGE" then
		EMReadScreen select_quarter, 1, 8, 14
		EMReadScreen IEVS_year, 4, 8, 22
	'ELSEIF match_type = "UBEN" THEN
	'	EMReadScreen IEVS_month, 2, 5, 68
	'	EMReadScreen IEVS_year, 4, 8, 71
	ELSEIF match_type = "BEER" or match_type = "UNVI" THEN
		EMReadScreen IEVS_year, 2, 8, 15
		IEVS_year = "20" & IEVS_year
		select_quarter = "YEAR"
	END IF
END IF

'------------------------------------------setting up case note header'
IF match_type = "BEER" THEN match_type_letter = "B"
IF match_type = "UBEN" THEN match_type_letter = "U"
IF match_type = "UNVI" THEN match_type_letter = "U"

'--------------------------------------------------------------------Client name
EmReadScreen panel_name, 4, 02, 52
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
'IF sent_date = "" THEN sent_date = replace(sent_date, " ", "/")
IF sent_date <> "" THEN sent_date = replace(sent_date, " ", "/")

'-----------------------------------------------------------------------------------------Initial dialog and do...loop
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 181, 240, "ATR Received"
  EditBox 55, 5, 55, 15, MAXIS_case_number
  EditBox 155, 5, 20, 15, MEMB_Number
  EditBox 55, 25, 55, 15, date_received
  DropListBox 85, 45, 55, 15, "Select One:"+chr(9)+"1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"YEAR", select_quarter
  DropListBox 85, 65, 55, 15, "Select One:"+chr(9)+"WAGE"+chr(9)+"BEER"+chr(9)+"UBEN"+chr(9)+"UNVI", match_type
  DropListBox 85, 85, 90, 15, "Select One:"+chr(9)+"MAIL"+chr(9)+"FAX"+chr(9)+"RCVD VERIFICATION", ATR_sent
  DropListBox 85, 105, 90, 15, "Select One:"+chr(9)+"DELETED DISQ"+chr(9)+"PENDING VERF"+chr(9)+"N/A", DISQ_action
  EditBox 65, 130, 110, 15, income_source
  EditBox 65, 150, 110, 15, source_address
  EditBox 65, 170, 110, 15, source_phone
  EditBox 65, 190, 110, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 65, 220, 50, 15
    CancelButton 125, 220, 50, 15
  Text 5, 10, 50, 10, "Case Number: "
  Text 120, 10, 30, 10, "MEMB #"
  Text 5, 30, 50, 10, "Date received:"
  Text 5, 50, 75, 10, "Match Period (quarter)"
  Text 5, 70, 65, 10, "Wage or Non-Wage"
  Text 45, 90, 30, 10, "ATR status"
  Text 5, 110, 75, 10, "DISQ panel addressed"
  Text 10, 135, 50, 10, "Source Name:"
  Text 30, 155, 30, 10, "Address:"
  Text 15, 175, 45, 10, "Fax or Phone:"
  Text 20, 195, 45, 10, "Other Notes:"
EndDialog

DO
    DO
        err_msg = ""
    	Dialog Dialog1
    	cancel_without_confirmation
    	IF IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
    	IF select_quarter = "Select One:" THEN err_msg = err_msg & vbNewLine & "Please select a quarter for the match"
    	IF match_type = "Select One:" THEN err_msg = err_msg & vbNewLine & "Please select a match type"
    	IF ATR_sent = "Select One:" THEN err_msg = err_msg & vbNewLine & "Please select how ATR was sent"
    	IF DISQ_action = "Select One:" THEN err_msg = err_msg & vbNewLine & "Please advise if DISQ panel was updated"
    	IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    LOOP UNTIL err_msg = ""
    CALL check_for_password_without_transmit(are_we_passworded_out)
LOOP UNTIL are_we_passworded_out = false

'--------------------------------------------------------------------sending the notice in IULA
EMwritescreen "005", 12, 46 'writing the resolve time to read for later
EMwritescreen "Y", 15, 37 'send Notice
transmit 'this will take us to IULB'\
ROW = 8
EMReadScreen IULB_first_line, 1, row, 6
IF IULB_first_line = "" THEN
	EMwritescreen "ATR RECEIVED " & date_received, row, 6
ELSE
	ROW = 9
	CALL clear_line_of_text(row, 6)
	EMwritescreen "ATR RECEIVED " & date_received, row, 6
END IF

'msgbox "Responded to difference notice has been updated"
TRANSMIT 'exiting IULA, helps prevent errors when going to the case note
''--------------------------------------------------------------------The case note & case note related code

diff_date = replace(diff_date, " ", "/")
IEVS_period = trim(IEVS_period)
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

'----------------------------------------------------------------the case note
CALL start_a_blank_case_note
IF match_type = "WAGE" THEN CALL write_variable_in_case_note("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH (" & first_name & ") ATR RECEIVED-----")
IF match_type = "BEER" or match_type = "UNVI" THEN CALL write_variable_in_case_note("-----" & IEVS_year & " NON-WAGE MATCH (" & first_name & ") " & "(" & match_type_letter & ") ATR RECEIVED-----")
IF match_type = "UBEN" THEN CALL write_variable_in_case_note("-----" & IEVS_period & " NON-WAGE MATCH (" & first_name & ") " & "(" & match_type_letter & ") ATR RECEIVED-----")
CALL write_bullet_and_variable_in_CASE_NOTE("Period", IEVS_period)
CALL write_bullet_and_variable_in_CASE_NOTE("Active Programs", programs)
CALL write_variable_in_CASE_NOTE("* Source information: " & source_income & income_source & "  " & source_address)
CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- -----")
CALL write_variable_in_CASE_NOTE("* Date ATR received: " & date_received)
IF DISQ_action = "DELETED DISQ" THEN CALL write_variable_in_CASE_NOTE("* Updated DISQ panel")
IF DISQ_action = "PENDING VERF" THEN CALL write_variable_in_CASE_NOTE("* Pending verification of income or asset")
CALL write_variable_in_CASE_NOTE("* IEVP updated as responded to difference notice - YES ")
IF ATR_sent <> "RCVD VERIFICATION" THEN
	CALL write_variable_in_CASE_NOTE("* Sent via: " & ATR_sent & " " & source_phone)
	CALL write_bullet_and_variable_in_case_note("Due Date", Due_date)
	CALL write_variable_in_CASE_NOTE("---DEU WILL PROCESS WHEN EMPLOYMENT VERIFICATION IS RETURNED. TEAM CAN REINSTATE CASE IF ALL NECESSARY PAPERWORK TO REINSTATE HAS BEEN RECEIVED---")
ELSE
	CALL write_variable_in_CASE_NOTE("---TEAM CAN REIN CASE IF ALL NECESSARY PAPERWORK TO REIN HAS BEEN RCVD---")
END IF
CALL write_bullet_and_variable_in_case_note("Other Notes", other_notes)
CALL write_variable_in_CASE_NOTE ("----- ----- ----- ----- -----")
CALL write_variable_in_CASE_NOTE ("DEBT ESTABLISHMENT UNIT 612-348-4290 EXT 1-1-1")
script_end_procedure_with_error_report("ATR case note updated successfully." & vbNewLine & "Please remember to update/delete the DISQ panel")
