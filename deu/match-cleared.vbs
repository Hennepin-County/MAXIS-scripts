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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: CALL changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
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
EMConnect ""
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
		 '----------------------------------------------------------------------------------------------------IEVP
		'Navigating deeper into the match interface
		CALL write_value_and_transmit("I", 6, 3)   		'navigates to INFC
		CALL write_value_and_transmit("IEVP", 20, 71)   'navigates to IEVP
		TRANSMIT
	    EMReadScreen err_msg, 7, 24, 2
	    IF err_msg = "NO IEVS" THEN script_end_procedure_with_error_report("An error occurred in IEVP, please process manually.")'checking for error msg'
	END IF
END IF

IF dail_check <> "DAIL" or match_type <> "WAGE" or match_type <> "BEER" or match_type <> "UBEN" or match_type <> "UNVI" or match_found = FALSE THEN
    CALL MAXIS_case_number_finder (MAXIS_case_number)
    MEMB_number = "01"
    BeginDialog case_number_dialog, 0, 0, 131, 65, "Case Number to clear match"
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
    		Dialog case_number_dialog
    		IF ButtonPressed = 0 THEN StopScript
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
	IF trim(IEVS_period) = "" THEN script_end_procedure_with_error_report("A match for the selected period could not be found. The script will now end.")
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
	ELSEIF match_type = "UBEN" THEN
		EMReadScreen IEVS_month, 2, 5, 68
		EMReadScreen IEVS_year, 4, 8, 71
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
'----------------------------------------------------------------------------dialogs
BeginDialog notice_action_dialog, 0, 0, 166, 90, "SEND DIFFERENCE NOTICE?"
	CheckBox 25, 35, 105, 10, "YES - Send Difference Notice", send_notice_checkbox
	CheckBox 25, 50, 130, 10, "NO - Continue Match Action to Clear", clear_action_checkbox
  Text 10, 10, 145, 20, "A difference notice has not been sent, would you like to send the difference notice now?"
  ButtonGroup ButtonPressed
    OkButton 60, 70, 45, 15
    CancelButton 110, 70, 45, 15
EndDialog

BeginDialog send_notice_dialog, 0, 0, 296, 160, "WAGE MATCH SEND DIFFERENCE NOTICE"
  CheckBox 10, 80, 70, 10, "Difference Notice", Diff_Notice_Checkbox
  CheckBox 110, 80, 90, 10, "Employment Verification", empl_verf_checkbox
  CheckBox 10, 95, 90, 10, "Authorization to Release", ATR_Verf_CheckBox
  CheckBox 110, 95, 80, 10, "Other (please specify)", other_checkbox
  EditBox 50, 120, 240, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 195, 140, 45, 15
    CancelButton 245, 140, 45, 15
  GroupBox 5, 5, 285, 55, "WAGE MATCH"
  GroupBox 5, 65, 200, 50, "Verification Requested: "
  Text 10, 20, 110, 10, "Case number: "  & MAXIS_case_number
  Text 120, 20, 165, 10, "Client name: "  & client_name
  Text 10, 40, 105, 10, "Active Programs: "  & programs
  Text 120, 40, 165, 15, "Income source: "   & income_source
  Text 5, 125, 40, 10, "Other notes: "
EndDialog

IF notice_sent = "N" THEN
	DO
    	err_msg = ""
    	Dialog notice_action_dialog
    	IF ButtonPressed = 0 THEN StopScript
    	IF send_notice_checkbox = UNCHECKED AND clear_action_checkbox = UNCHECKED THEN err_msg = err_msg & vbNewLine & "* Please select an answer to continue."
    	IF send_notice_checkbox = CHECKED AND clear_action_checkbox = CHECKED THEN err_msg = err_msg & vbNewLine & "* Please select only one answer to continue."
    	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
END IF
CALL check_for_password_without_transmit(are_we_passworded_out)

IF send_notice_checkbox = CHECKED THEN
'----------------------------------------------------------------Defaulting checkboxes to being checked (per DEU instruction)
    Diff_Notice_Checkbox = CHECKED
    ATR_Verf_CheckBox = CHECKED
'---------------------------------------------------------------------send notice dialog and dialog DO...loop
	DO
    	err_msg = ""
    	Dialog send_notice_dialog
    	cancel_confirmation
		IF other_checkbox = CHECKED and other_notes = "" THEN err_msg = err_msg & vbNewLine & "* Please specify what other is to continue."
    	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
    CALL check_for_password_without_transmit(are_we_passworded_out)
'--------------------------------------------------------------------sending the notice in IULA
    EMwritescreen "005", 12, 46 'writing the resolve time to read for later
    EMwritescreen "Y", 14, 37 'send Notice
	'msgbox "Difference Notice Sent"
	TRANSMIT 'goes into IULA
	'removed the IULB information '
	TRANSMIT'exiting IULA, helps prevent errors when going to the case note
    '-----------------------------------------------------------------------------------Claim Referral Tracking
    action_date = date & ""

    '-----------------------------------------------------------------Going to the MISC panel
	'Going to the MISC panel to add claim referral tracking information
    Call navigate_to_MAXIS_screen ("STAT", "MISC")
    Row = 6
    EmReadScreen panel_number, 1, 02, 73
    If panel_number = "0" then
    	EMWriteScreen "NN", 20,79
    	TRANSMIT
		'CHECKING FOR MAXIS PROGRAMS ARE INACTIVE'
		EmReadScreen MISC_error_msg,  74, 24, 02
		IF trim(MISC_error_msg) = "" THEN
	        case_note_only = FALSE
		else
			maxis_error_check = MsgBox("*** NOTICE!!!***" & vbNewLine & "Continue to case note only?" & vbNewLine & MISC_error_msg & vbNewLine, vbYesNo + vbQuestion, "Message handling")
			IF maxis_error_check = vbYes THEN
				case_note_only = TRUE 'this will case note only'
			END IF
			IF maxis_error_check= vbNo THEN
				case_note_only = FALSE 'this will update the panels and case note'
			END IF
		END IF
    ELSE
		IF case_note_only = FALSE THEN
			Do
    			'Checking to see if the MISC panel is empty, if not it will find a new line'
    			EmReadScreen MISC_description, 25, row, 30
    			MISC_description = replace(MISC_description, "_", "")
    			If trim(MISC_description) = "" then
    				PF9
    				EXIT DO
    			Else
    				row = row + 1
    			End if
    		Loop Until row = 17
    		If row = 17 then MsgBox("There is not a blank field in the MISC panel. Please delete a line(s), and run script again or update manually.")
    	End if
		'writing in the action taken and date to the MISC panel
		EMWriteScreen "Claim Determination", Row, 30
		EMWriteScreen date, Row, 66
		PF3
	END IF 'checking to make sure maxis case is active
	EMWriteScreen "Initial Claim Referral", Row, 30
    EMWriteScreen date, Row, 66
    PF3

    'The case note-------------------------------------------------------------------------------------------------
    start_a_blank_case_note
    Call write_variable_in_case_note("-----Claim Referral Tracking - Initial Claim Referral-----")
	IF case_note_only = TRUE THEN Call write_variable_in_case_note("Maxis case is inactive unable to add or update MISC panel")
    Call write_bullet_and_variable_in_case_note("Action Date", action_date)
    Call write_bullet_and_variable_in_case_note("Active Program(s)", programs)
    IF next_action = "Sent Request for Additional Info" THEN CALL write_variable_in_case_note("* Additional verifications requested, TIKL set for 10 day return.")
    Call write_bullet_and_variable_in_case_note("Other Notes", other_notes)
    Call write_variable_in_case_note("* Entries for these potential claims must be retained until further notice.")
    Call write_variable_in_case_note("-----")
    Call write_variable_in_case_note(worker_signature)
    PF3
	'--------------------------------------------------------------------The case note & case note related code
	pending_verifs = ""
  	IF Diff_Notice_Checkbox = CHECKED THEN pending_verifs = pending_verifs & "Difference Notice, "
	IF empl_verf_checkbox = CHECKED THEN pending_verifs = pending_verifs & "EVF, "
	IF ATR_Verf_CheckBox = CHECKED THEN pending_verifs = pending_verifs & "ATR, "
	IF other_checkbox = CHECKED THEN pending_verifs = pending_verifs & "Other, "

    '-------------------------------------------------------------------trims excess spaces of pending_verifs
  	pending_verifs = trim(pending_verifs) 	'takes the last comma off of pending_verifs when autofilled into dialog if more than one app date is found and additional app is selected
	IF right(pending_verifs, 1) = "," THEN pending_verifs = left(pending_verifs, len(pending_verifs) - 1)
	IF match_type = "WAGE" THEN
		IF select_quarter = 1 THEN IEVS_quarter = "1ST"
		IF select_quarter = 2 THEN IEVS_quarter = "2ND"
		IF select_quarter = 3 THEN IEVS_quarter = "3RD"
		IF select_quarter = 4 THEN IEVS_quarter = "4TH"
	END IF
	IEVS_period = replace(IEVS_period, "/", " to ")
	Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days

	'---------------------------------------------------------------------DIFF NOTC case note
  	start_a_blank_case_note
	IF match_type = "WAGE" THEN CALL write_variable_in_case_note("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH(" & first_name & ") DIFF NOTICE SENT-----")
	IF match_type = "BEER" or match_type = "UNVI" THEN CALL write_variable_in_case_note("-----" & IEVS_year & " NON-WAGE MATCH(" & match_type_letter & ") " & "(" & first_name & ") DIFF NOTICE SENT-----")
	IF match_type = "UBEN" THEN CALL write_variable_in_case_note("-----" & IEVS_month & " NON-WAGE MATCH(" & match_type_letter & ") " & "(" & first_name & ") DIFF NOTICE SENT-----")
	CALL write_bullet_and_variable_in_case_note("Client Name", client_name)
	CALL write_bullet_and_variable_in_case_note("Period", IEVS_period)
  	CALL write_bullet_and_variable_in_case_note("Active Programs", programs)
	CALL write_bullet_and_variable_in_case_note("Source of income", income_source)
	CALL write_variable_in_case_note ("----- ----- -----")
  	CALL write_bullet_and_variable_in_case_note("Verification Requested", pending_verifs)
  	CALL write_bullet_and_variable_in_case_note("Verification Due", Due_date)
	CALL write_variable_in_case_note ("* Client must be provided 10 days to return requested verifications *")
  	CALL write_bullet_and_variable_in_case_note("Other notes", other_notes)
  	CALL write_variable_in_case_note("----- ----- ----- ----- ----- ----- -----")
  	CALL write_variable_in_case_note ("DEBT ESTABLISHMENT UNIT 612-348-4290 EXT 1-1-1")
END IF

IF clear_action_checkbox = CHECKED or notice_sent = "Y" THEN
	IF sent_date <> "" THEN MsgBox("A difference notice was sent on " & sent_date & "." & vbNewLine & "The script will now navigate to clear the match.")
	'date_received = date & ""

	BeginDialog cleared_match_dialog, 0, 0, 316, 175, "MATCH CLEARED"
	  EditBox 55, 5, 40, 15, MAXIS_case_number
	  EditBox 165, 5, 20, 15, MEMB_number
	  DropListBox 250, 5, 60, 15, "Select One:"+chr(9)+"BEER"+chr(9)+"BNDX"+chr(9)+"SDXS/ SDXI"+chr(9)+"UNVI"+chr(9)+"UBEN"+chr(9)+"WAGE", match_type
	  DropListBox 55, 25, 40, 15, "Select One:"+chr(9)+"1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"YEAR"+chr(9)+"N/A", select_quarter
	  EditBox 165, 25, 20, 15, resolve_time
	  CheckBox 205, 35, 90, 10, "Authorization to release", ATR_Verf_CheckBox
	  CheckBox 205, 45, 70, 10, "Difference notice", Diff_Notice_Checkbox
	  CheckBox 205, 55, 90, 10, "Employment verification", empl_verf_checkbox
	  CheckBox 205, 65, 70, 10, "School verification", school_verf_checkbox
	  CheckBox 205, 75, 80, 10, "Other (please specify)", other_checkbox
	  DropListBox 70, 45, 115, 15, "Select One:"+chr(9)+"BC - Case Closed"+chr(9)+"BN - Already known, No Savings"+chr(9)+"BE - Child"+chr(9)+"BE - No Change"+chr(9)+"BE - NC Non-collectible"+chr(9)+"BE - OP Entered"+chr(9)+"BO - Other"+chr(9)+"BP - Wrong Person"+chr(9)+"CC - Claim Entered"+chr(9)+"NC - Non Cooperation", resolution_status
	  DropListBox 120, 65, 65, 15, "Select One:"+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"N/A", change_response
	  DropListBox 120, 85, 65, 15, "Select One:"+chr(9)+"DISQ Deleted"+chr(9)+"Pending Verif"+chr(9)+"No"+chr(9)+"N/A", DISQ_action
	  EditBox 270, 95, 40, 15, date_received
	  EditBox 270, 115, 40, 15, exp_grad_date
	  EditBox 55, 135, 255, 15, other_notes
	  ButtonGroup ButtonPressed
	    OkButton 215, 155, 45, 15
	    CancelButton 265, 155, 45, 15
	  CheckBox 10, 105, 115, 10, "Check here if 10 day has passed", TIKL_checkbox
	  CheckBox 10, 120, 150, 10, "Check here to run Claim Referral Tracking", claim_referral
	  Text 5, 10, 50, 10, "Case number: "
	  Text 110, 10, 50, 10, "MEMB number:"
	  Text 200, 10, 40, 10, "Match Type: "
	  Text 5, 30, 50, 10, "Match period: "
	  Text 100, 30, 65, 10, "Resolve time (min): "
	  Text 5, 50, 60, 10, "Resolution status: "
	  Text 5, 70, 105, 10, "Responded to difference notice: "
	  GroupBox 195, 25, 115, 65, "Verification Used to Clear: "
	  Text 190, 100, 75, 10, "Date verif rcvd/on file:"
	  Text 200, 120, 65, 10, "Expected grad date:"
	  Text 10, 140, 40, 10, "Other notes: "
	  Text 35, 90, 75, 10, "DISQ panel addressed:"
	EndDialog
	DO
		err_msg = ""
		Dialog cleared_match_dialog
		cancel_confirmation
		IF IsNumeric(resolve_time) = false or len(resolve_time) > 3 THEN err_msg = err_msg & vbNewLine & "* Enter a valid numeric resolved time, ie 005."
		IF resolve_time = "" THEN err_msg = err_msg & vbNewLine & "Please complete resolve time."
		IF date_received = "" THEN
			IF resolution_status <> "BN - Already known, No Savings" THEN err_msg = err_msg & vbNewLine & "Please advise of date verification was recieved in ECF."
		END IF
		IF date_received = "" THEN
			IF resolution_status <> "NC - Non Cooperation" THEN err_msg = err_msg & vbNewLine & "Please advise of date verification was recieved in ECF."
		END IF
		IF other_checkbox = CHECKED and other_notes = "" THEN err_msg = err_msg & vbNewLine & "Please advise what other verification was used to clear the match."
		IF change_response = "Select One:" THEN err_msg = err_msg & vbNewLine & "Did the client respond to Difference Notice?"
		IF resolution_status = "Select One:" THEN err_msg = err_msg & vbNewLine & "Please select a resolution status to continue."
		IF resolution_status = "BE - No Change" AND other_notes = "" THEN err_msg = err_msg & vbNewLine & "When clearing using BE other notes must be completed."
		IF resolution_status = "BE - Child" AND exp_grad_date = "" THEN err_msg = err_msg & vbNewLine & "When clearing using BE - Child graduation date and date rcvd must be completed."
		If resolution_status = "CC - Claim Entered" AND programs = "Health Care" or programs = "Medical Assistance" THEN err_msg = err_msg & vbNewLine & "* System does not allow HC or MA cases to be cleared with the code 'CC - Claim Entered'."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password_without_transmit(are_we_passworded_out)

	'----------------------------------------------------------------------------------------------------RESOLVING THE MATCH
	EMWriteScreen resolve_time, 12, 46	    'resolved notes depending on the resolution_status
	IF resolution_status = "BC - Case Closed" THEN rez_status = "BC"
	IF resolution_status = "BE - Child" THEN rez_status = "BE"
	IF resolution_status = "BE - No Change" THEN rez_status = "BE"
	IF resolution_status = "BE - OP Entered" THEN rez_status = "BE"
	IF resolution_status = "BE - NC Non-collectible" THEN rez_status = "BE"
	IF resolution_status = "BN - Already known, No Savings" THEN rez_status = "BN"
	IF resolution_status = "BO - Other" THEN rez_status = "BO"
	IF resolution_status = "BP - Wrong Person"  THEN rez_status = "BP"
	IF resolution_status = "CC - Claim Entered" THEN rez_status = "CC"
	IF resolution_status = "NC - Non Cooperation" THEN rez_status = "NC"
	'CC cannot be used - ACTION CODE FOR ACTH OR ACTM IS INVALID
	'checked these all to programS'
	EMwritescreen rez_status, 12, 58
	IF change_response = "YES" THEN
		EMwritescreen "Y", 15, 37
	ELSE
		EMwritescreen "N", 15, 37
	END IF
	TRANSMIT 'IULB
	'----------------------------------------------------------------------------------------writing the note on IULB
	IULB_error_msg = ""
	EMReadScreen IULB_error_msg, 50, 24, 2
	IULB_error_msg = trim(IULB_error_msg)
	IF IULB_error_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & IULB_error_msg & vbNewLine
	If IULB_error_msg = "ACTION CODE" THEN script_end_procedure_with_error_report(IULB_error_msg & vbNewLine & "Please ensure you are selecting the correct code for resolve. PF10 to ensure the match can be resolved using the script.")'checking for error msg'
	IF resolution_status = "BC - Case Closed" 	THEN EMWriteScreen "Case closed. " & other_notes, 8, 6   							'BC
	IF resolution_status = "BE - No Change" THEN EMWriteScreen "No change. " & other_notes, 8, 6 									'BE
	IF resolution_status = "BE - Child" THEN EMWriteScreen "No change, minor child income excluded. " & other_notes, 8, 6 			'BE - child
	IF resolution_status = "BE - OP Entered" THEN EMWriteScreen "OP entered other programs" & other_notes, 8, 6
	IF resolution_status = "BE - NC Non-collectible" THEN EMWriteScreen "Non-Coop remains, but claim is non-collectible ", 8, 6
	IF resolution_status = "BN - Already known, No Savings" THEN EMWriteScreen "Already known - No savings. " & other_notes, 8, 6 	'BN
	IF resolution_status = "BO - Other" THEN EMWriteScreen "HC Claim entered. " & other_notes, 8, 6 								'BO
	IF resolution_status = "BP - Wrong Person" THEN EMWriteScreen "Client name and wage earner name are different. " & other_notes, 8, 6
	IF resolution_status = "CC - Claim Entered" THEN EMWriteScreen "Claim entered. " & other_notes, 8, 6 						 	'CC
	IF resolution_status = "NC - Non Cooperation" THEN EMWriteScreen "Non-coop, requested verf not in ECF, " & other_notes, 8, 6 	'NC
	'msgbox "did the notes input?"
	TRANSMIT 'this will take us back to IEVP main menu'
	'------------------------------------------------------------------back on the IEVP menu, making sure that the match cleared
	EMReadScreen days_pending, 5, 7, 72
	days_pending = trim(days_pending)
	IF IsNumeric(days_pending) = TRUE THEN
		match_cleared = FALSE
		script_end_procedure_with_error_report("This match did not appear to clear. Please check case, and try again.")
	ELSE
	  match_cleared = TRUE
	END IF

  	IF match_type = "WAGE" THEN
  	'Updated IEVS_periodto write into case note
  		IF select_quarter = 1 THEN IEVS_quarter = "1ST"
  		IF select_quarter = 2 THEN IEVS_quarter = "2ND"
  		IF select_quarter = 3 THEN IEVS_quarter = "3RD"
  		IF select_quarter = 4 THEN IEVS_quarter = "4TH"
  	END IF

	IEVS_period = replace(IEVS_period, "/", " to ")
	Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days requested for HEADER of casenote'
	PF3 'back to the DAIL'
   '----------------------------------------------------------------the case match CLEARED note
	start_a_blank_case_note
	IF resolution_status <> "NC - Non Cooperation" THEN
		IF match_type = "WAGE" THEN CALL write_variable_in_case_note("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH (" & first_name & ") CLEARED " & rez_status & "-----")
		IF match_type = "BEER" or match_type = "UNVI" THEN CALL write_variable_in_case_note("-----" & IEVS_year & " NON-WAGE MATCH (" & match_type_letter & ") " & "(" & first_name & ") CLEARED " & rez_status & "-----")
		IF match_type = "UBEN" THEN CALL write_variable_in_case_note("-----" & IEVS_month & " NON-WAGE MATCH (" & match_type_letter & ") " & "(" & first_name & ") CLEARED " & rez_status & "-----")
		CALL write_bullet_and_variable_in_case_note("Period", IEVS_period)
		CALL write_bullet_and_variable_in_case_note("Active Programs", programs)
		CALL write_bullet_and_variable_in_case_note("Source of income", income_source)
		CALL write_variable_in_case_note ("----- ----- -----")
		CALL write_bullet_and_variable_in_case_note("Date verification received in ECF", date_received)
		IF resolution_status = "BC - Case Closed" THEN CALL write_variable_in_case_note("* Case closed. ")
		IF resolution_status = "BE - Child" THEN
			CALL write_variable_in_case_note("* Income is excluded for minor child in school.")
			CALL write_bullet_and_variable_in_case_note("Expected graduation date", exp_grad_date)
		END IF
		IF resolution_status = "BE - OP Entered" THEN CALL write_variable_in_case_note("* Overpayments or savings were found related to this match.")
		IF resolution_status = "BE - No Change" THEN CALL write_variable_in_case_note("* No Overpayments or savings were found related to this match.")
		IF resolution_status = "BE - NC Non-collectible" THEN CALL write_variable_in_case_note("* No collectible overpayments or savings were found related to this match. Client is still non-coop.")
		IF resolution_status = "BN - Already known, No Savings" THEN CALL write_variable_in_case_note("* Client reported income. Correct income is in JOBS/BUSI and budgeted.")
		IF resolution_status = "BO - Other" THEN CALL write_variable_in_case_note("* HC Claim entered. ")
		IF resolution_status = "BP - Wrong Person" THEN CALL write_variable_in_case_note("* Client name and wage earner name are different.  Client's SSN has been verified. No overpayment or savings related to this match.")
		IF resolution_status = "CC - Claim Entered" THEN CALL write_variable_in_case_note("* Claim entered.")
  		IF change_response <> "N/A" THEN CALL write_bullet_and_variable_in_case_note("Responded to Difference Notice", change_response)
  		CALL write_bullet_and_variable_in_case_note("Other notes", other_notes)
  		CALL write_variable_in_case_note("----- ----- ----- ----- -----")
  		CALL write_variable_in_case_note ("DEBT ESTABLISHMENT UNIT 612-348-4290 EXT 1-1-1")
		PF3
	ELSEIF resolution_status = "NC - Non Cooperation" THEN   'Navigates to TIKL
		IF match_type = "WAGE" THEN CALL write_variable_in_case_note("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH (" & first_name & ") NON-COOPERATION-----")
		IF match_type = "BEER" or match_type = "UNVI" THEN CALL write_variable_in_case_note("-----" & IEVS_year & " NON-WAGE MATCH (" & match_type_letter & ") " & "(" & first_name & ") NON-COOPERATION-----")
		IF match_type = "UBEN" THEN CALL write_variable_in_case_note("-----" & IEVS_month & " NON-WAGE MATCH (" & match_type_letter & ") " & "(" & first_name & ") NON-COOPERATION-----")
		CALL write_bullet_and_variable_in_case_note("Period", IEVS_period)
		CALL write_bullet_and_variable_in_case_note("Active Programs", programs)
		CALL write_bullet_and_variable_in_case_note("Source of income", income_source)
		CALL write_variable_in_case_note ("----- ----- -----")
		CALL write_variable_in_case_note("* Client failed to cooperate wth wage match.")
		IF DISQ_action <> "Select One:" THEN CALL write_bullet_and_variable_in_case_note("STAT/DISQ addressed for each program", DISQ_action)
		CALL write_bullet_and_variable_in_case_note("Date Diff notice sent", sent_date)
		CALL write_variable_in_case_note("* Case approved to close.")
		CALL write_variable_in_case_note("* Client needs to provide: ATR, Income Verification, Difference Notice.")
		IF change_response <> "N/A" THEN CALL write_bullet_and_variable_in_case_note("Responded to Difference Notice", change_response)
		CALL write_bullet_and_variable_in_case_note("* Other notes", other_notes)
		CALL write_variable_in_case_note("----- ----- ----- ----- -----")
		CALL write_variable_in_case_note ("DEBT ESTABLISHMENT UNIT 612-348-4290 EXT 1-1-1")
		PF3

		'-------------------------------The following will generate a TIKL formatted date for 10 days from now, and add it to the TIKL
		Call navigate_to_MAXIS_screen("DAIL", "WRIT")
	  	CALL create_MAXIS_friendly_date(date, 10, 5, 18)
	 	CALL write_variable_in_TIKL("CLOSE FOR NON-COOP, CREATE DISQ(S) FOR " & first_name)
	 	PF3		'Exits and saves TIKL
		'script_end_procedure("Success! Updated match, and a TIKL created.")
		script_end_procedure_with_error_report("Success! Updated match, and a TIKL created.")
	END IF
	script_end_procedure_with_error_report("Match has been acted on. Please take any additional action needed for your case.")
END IF
