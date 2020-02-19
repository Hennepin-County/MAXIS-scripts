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
CALL changelog_update("01/31/2020", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'---------------------------------------------------------------------THE SCRIPT
EMConnect ""
CALL MAXIS_case_number_finder (MAXIS_case_number)
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 131, 65, "Case Number"
  EditBox 65, 5, 60, 15, MAXIS_case_number
  EditBox 65, 25, 20, 15, MEMB_number
  ButtonGroup ButtonPressed
    OkButton 30, 45, 45, 15
    CancelButton 80, 45, 45, 15
  Text 5, 30, 60, 10, "Member Number:"
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
'---------------------------------------------------------------------Reading potential errors for out-of-county cases
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
	ELSEIF match_type = "BEER" THEN
		EMReadScreen IEVS_year, 2, 8, 15
		IEVS_year = "20" & IEVS_year
	ELSEIF match_type = "UNVI" THEN
		EMReadScreen IEVS_year, 4, 8, 15
		select_quarter = "YEAR"
	END IF
END IF

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

IF notice_sent = "N" THEN
    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 166, 90, "SEND DIFFERENCE NOTICE?"
    	CheckBox 25, 35, 105, 10, "YES - Send Difference Notice", send_notice_checkbox
    	CheckBox 25, 50, 130, 10, "NO - Continue Match Action to Clear", clear_action_checkbox
      Text 10, 10, 145, 20, "A difference notice has not been sent, would you like to send the difference notice now?"
      ButtonGroup ButtonPressed
    	OkButton 60, 70, 45, 15
    	CancelButton 110, 70, 45, 15
    EndDialog
	DO
    	err_msg = ""
    	Dialog Dialog1
    	cancel_without_confirmation
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
	'-------------------------------------------------------------------------------------------------DIALOG
	Dialog1 = "" 'Blanking out previous dialog detail
	BeginDialog Dialog1, 0, 0, 296, 160, "MATCH SEND DIFFERENCE NOTICE"
	  CheckBox 10, 80, 70, 10, "Difference Notice", Diff_Notice_Checkbox
	  CheckBox 10, 95, 90, 10, "Authorization to Release", ATR_Verf_CheckBox
	  CheckBox 110, 80, 80, 10, "Lottery/Gaming Form", lottery_verf_checkbox
	  CheckBox 110, 95, 80, 10, "Rental Income Form", rental_checkbox
	  CheckBox 200, 80, 90, 10, "Employment Verification", empl_verf_checkbox
	  CheckBox 200, 95, 80, 10, "Other (please specify)", other_checkbox
	  EditBox 50, 120, 240, 15, other_notes
	  CheckBox 5, 145, 180, 10, "Check to add claim referral tracking(SNAP and MF)", claim_referral_tracking_checkbox
	  ButtonGroup ButtonPressed
	    OkButton 195, 140, 45, 15
	    CancelButton 245, 140, 45, 15
	  Text 10, 20, 110, 10, "Case number: "   & MAXIS_case_number
	  Text 120, 20, 165, 10, "Client name: "  & client_name
	  Text 10, 40, 105, 10, "Active Programs: "  & programs
	  Text 120, 40, 165, 15, "Income source: "   & income_source
	  GroupBox 5, 65, 285, 50, "Verification Requested: "
	  Text 5, 125, 40, 10, "Other notes:"
	  GroupBox 5, 5, 285, 55, "WAGE MATCH"
	EndDialog
	'---------------------------------------------------------------Defaulting checkboxes to being checked (per DEU instruction)
    Diff_Notice_Checkbox = CHECKED
    ATR_Verf_CheckBox = CHECKED
    '---------------------------------------------------------------------send notice dialog and dialog DO...loop
	DO
    	err_msg = ""
    	Dialog Dialog1
    	cancel_without_confirmation
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
	IF claim_referral_tracking_checkbox = CHECKED THEN
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
	    EMWriteScreen "Initial Claim Referral", Row, 30
	    EMWriteScreen date, Row, 66
	    TRANSMIT

        'The case note-------------------------------------------------------------------------------------------------
        start_a_blank_case_note
        Call write_variable_in_case_note("-----Claim Referral Tracking - Initial Claim Referral-----")
	    Call write_bullet_and_variable_in_case_note("Action Date", action_date)
        Call write_bullet_and_variable_in_case_note("Active Program(s)", programs)
        IF next_action = "Sent Request for Additional Info" THEN CALL write_variable_in_case_note("* Additional verifications requested, follow up set for 10 day return.")
        Call write_bullet_and_variable_in_case_note("Other Notes", other_notes)
        Call write_variable_in_case_note("* Entries for these potential claims must be retained until further notice.")
	    IF case_note_only = TRUE THEN Call write_variable_in_case_note("Maxis case is inactive unable to add or update MISC panel")
        Call write_variable_in_case_note("-----")
        Call write_variable_in_case_note(worker_signature)
        PF3
	END IF
	'--------------------------------------------------------------------The case note & case note related code
	pending_verifs = ""
  	IF Diff_Notice_Checkbox = CHECKED THEN pending_verifs = pending_verifs & "Difference Notice, "
	IF empl_verf_checkbox = CHECKED THEN pending_verifs = pending_verifs & "EVF, "
	IF ATR_Verf_CheckBox = CHECKED THEN pending_verifs = pending_verifs & "ATR, "
	IF lottery_verf_checkbox = CHECKED THEN pending_verifs = pending_verifs & "Lottery/Gaming Form, "
	IF rental_checkbox =  CHECKED THEN pending_verifs = pending_verifs & "Rental Income Form, "
	IF other_checkbox = CHECKED THEN pending_verifs = pending_verifs & "Other, "

    '-------------------------------------------------------------------trims excess spaces of pending_verifs
    pending_verifs = trim(pending_verifs) 	'takes the last comma off of pending_verifs when autofilled into dialog if more more than one app date is found and additional app is selected
	IF right(pending_verifs, 1) = "," THEN pending_verifs = left(pending_verifs, len(pending_verifs) - 1)
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

	'---------------------------------------------------------------------DIFF NOTC case note
  	start_a_blank_case_note
	IF match_type = "WAGE" THEN CALL write_variable_in_case_note("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH (" & first_name & ") DIFF NOTICE SENT-----")
	IF match_type = "BEER" THEN CALL write_variable_in_case_note("-----" & IEVS_year & " NON-WAGE MATCH (" & first_name & ") (B) DIFF NOTICE SENT-----")
	IF match_type = "UNVI" THEN CALL write_variable_in_case_note("-----" & IEVS_year & " NON-WAGE MATCH (" & first_name & ") (U) DIFF NOTICE SENT-----")
	IF match_type = "UBEN" THEN CALL write_variable_in_case_note("-----" & IEVS_period & " NON-WAGE MATCH (" & first_name & ") (U) DIFF NOTICE SENT-----")
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
END IF'end of sending notice
'-------------------------------------------------------------------------------------end of sending notice

IF clear_action_checkbox = CHECKED or notice_sent = "Y" THEN
	IF sent_date <> "" THEN MsgBox("A difference notice was sent on " & sent_date & "." & vbNewLine & "The script will now navigate to clear the match.")
	'date_received = date & ""
	'-------------------------------------------------------------------------------------------------DIALOG
	Dialog1 = "" 'Blanking out previous dialog detail
	BeginDialog Dialog1, 0, 0, 321, 185, "MATCH CLEARED - CASE NUMBER: " & MAXIS_case_number
	  DropListBox 55, 5, 50, 15, "Select One:"+chr(9)+"BEER"+chr(9)+"BNDX"+chr(9)+"SDXS/SDXI"+chr(9)+"UNVI"+chr(9)+"UBEN"+chr(9)+"WAGE", match_type
	  EditBox 170, 5, 15, 15, MEMB_number
	  DropListBox 55, 25, 40, 15, "Select One:"+chr(9)+"1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"YEAR"+chr(9)+"N/A", select_quarter
	  EditBox 170, 25, 15, 15, resolve_time
	  DropListBox 70, 45, 110, 15, "Select One:"+chr(9)+"CB-Ovrpmt And Future Save"+chr(9)+"CC-Overpayment Only"+chr(9)+"CF-Future Save"+chr(9)+"CA-Excess Assets"+chr(9)+"CI-Benefit Increase"+chr(9)+"CP-Applicant Only Savings"+chr(9)+"BC-Case Closed"+chr(9)+"BE-Child"+chr(9)+"BE-No Change"+chr(9)+"BE-NC-Non-collectible"+chr(9)+"BE-Overpayment Entered"+chr(9)+"BN-Already Known-No Savings"+chr(9)+"BI-Interface Prob"+chr(9)+"BP-Wrong Person"+chr(9)+"BU-Unable To Verify"+chr(9)+"BO-Other"+chr(9)+"NC-Non Cooperation", resolution_status
	  DropListBox 120, 65, 60, 15, "Select One:"+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"N/A", change_response
	  DropListBox 120, 85, 60, 15, "Select One:"+chr(9)+"DISQ Added"+chr(9)+"DISQ Deleted"+chr(9)+"Pending Verif"+chr(9)+"No"+chr(9)+"N/A", DISQ_action
	  EditBox 270, 15, 40, 15, date_received
	  CheckBox 195, 35, 70, 10, "Difference Notice", Diff_Notice_Checkbox
	  CheckBox 195, 45, 90, 10, "Authorization to Release", ATR_Verf_CheckBox
	  CheckBox 195, 55, 90, 10, "Employment verification", EVF_checkbox
	  CheckBox 195, 65, 80, 10, "Lottery/Gaming Form", lottery_verf_checkbox
	  CheckBox 195, 75, 80, 10, "Rental Income Form", rental_checkbox
	  CheckBox 195, 85, 80, 10, "Other (please specify)", other_checkbox
	  EditBox 270, 100, 40, 15, exp_grad_date
	  CheckBox 5, 110, 175, 10, "Check here if 10 day has passed - TIKL will be set ", TIKL_checkbox
	  CheckBox 5, 120, 255, 10, "Check to update claim referral tracking(SNAP and MF) Overpayment Exists", overpayment_exists_checkbox
	  CheckBox 5, 130, 265, 10, "Check to update claim referral tracking(SNAP and MF) No Overpayment Exists", no_overpayment_checkbox
	  EditBox 50, 145, 265, 15, other_notes
	  ButtonGroup ButtonPressed
	    OkButton 220, 165, 45, 15
	    CancelButton 270, 165, 45, 15
	  Text 5, 10, 40, 10, "Match Type: "
	  Text 5, 30, 45, 10, "Match Period: "
	  Text 105, 30, 65, 10, "Resolve time (min): "
	  GroupBox 190, 5, 125, 115, "Verification Used to Clear: "
	  Text 5, 50, 60, 10, "Resolution Status: "
	  Text 5, 70, 110, 10, "Responded to Difference Notice: "
	  Text 40, 90, 75, 10, "DISQ panel addressed:"
	  Text 5, 150, 40, 10, "Other notes: "
	  Text 195, 20, 75, 10, "Date verif rcvd/on file:"
	  Text 110, 10, 55, 10, "MEMB Number:"
	  Text 195, 105, 65, 10, "Expected grad date:"
	EndDialog

	DO
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
		IF IsNumeric(resolve_time) = false or len(resolve_time) > 3 THEN err_msg = err_msg & vbNewLine & "Please enter a valid numeric resolved time, ie 005."
		'IF resolution_status = "CB-Ovrpmt And Future Save" or resolution_status = "CC-Overpayment Only" or resolution_status = "CF-Future Save" or resolution_status = "CA-Excess Assets" or resolution_status = "CI-Benefit Increase" or resolution_status = "CP-Applicant Only Savings" or resolution_status = "BC-Case Closed" or resolution_status = "BE-No Change" or resolution_status = "BE-NC-Non-collectible" or resolution_status = "BN-Already Known-No Savings" or resolution_status ="BP-Wrong Person" or resolution_status = "BO-Other" and date_received = "" THEN err_msg = err_msg & vbNewLine & "Please advise of date verification was recieved in ECF."
		IF other_checkbox = CHECKED and other_notes = "" THEN err_msg = err_msg & vbNewLine & "Please advise what other verification was used to clear the match."
		IF change_response = "Select One:" THEN err_msg = err_msg & vbNewLine & "Did the client respond to Difference Notice?"
		IF resolution_status = "Select One:" THEN err_msg = err_msg & vbNewLine & "Please select a resolution status to continue."
		IF resolution_status = "BE-No Change" AND other_notes = "" THEN err_msg = err_msg & vbNewLine & "When clearing using BE other notes must be completed."
		IF resolution_status = "BE-Child" AND exp_grad_date = "" THEN err_msg = err_msg & vbNewLine & "When clearing using BE - Child graduation date and date rcvd must be completed."
		If resolution_status = "CC-Overpayment Only" AND programs = "Health Care" or programs = "Medical Assistance" THEN err_msg = err_msg & vbNewLine & "System does not allow HC or MA cases to be cleared with the code 'CC - Claim Entered'."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password_without_transmit(are_we_passworded_out)
END IF
IF resolution_status = "CC-Overpayment Only" THEN
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 361, 280, "Match Cleared CC Claim Entered"
     EditBox 60, 5, 40, 15, MAXIS_case_number
     EditBox 140, 5, 20, 15, memb_number
      EditBox 230, 5, 20, 15, OT_resp_memb
      DropListBox 310, 5, 45, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", fraud_referral
      EditBox 60, 25, 40, 15, discovery_date
      DropListBox 210, 25, 40, 15, "Select:"+chr(9)+"BEER"+chr(9)+"NONE"+chr(9)+"UNVI"+chr(9)+"UBEN"+chr(9)+"WAGE", match_type
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
      CheckBox 70, 265, 130, 10, "Check if the EVF/ATR is still needed", ATR_needed_checkbox
    EndDialog
    Do
        Do
        	err_msg = ""
        	dialog Dialog1
        	cancel_without_confirmation
        	IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
        	IF select_quarter = "Select:" THEN err_msg = err_msg & vbnewline & "* You must select a match period entry-select other for UBEN."
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
        	IF match_type = "Select:" THEN err_msg = err_msg & vbnewline & "* You must select a match type entry."
        	IF EVF_used = "" then err_msg = err_msg & vbNewLine & "* Please enter verication used for the income recieved. If no verification was received enter N/A."
        	IF isdate(income_rcvd_date) = False or income_rcvd_date = "" then err_msg = err_msg & vbNewLine & "* Please enter a valid date for the income recieved."
        	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
        LOOP UNTIL err_msg = ""
        CALL check_for_password_without_transmit(are_we_passworded_out)
    Loop until are_we_passworded_out = false
    '----------------------------------------------------------------------------------------------------RESOLVING THE MATCH
        EmReadScreen panel_name, 4, 02, 52
        IF panel_name <> "IULA" THEN
            EmReadScreen back_panel_name, 4, 2, 52
            If back_panel_name <> "IEVP" Then
                CALL back_to_SELF
                CALL navigate_to_MAXIS_screen("INFC" , "____")
                CALL write_value_and_transmit("IEVP", 20, 71)
                CALL write_value_and_transmit(SSN_number_read, 3, 63)
            End If
            CALL write_value_and_transmit("U", row, 3)   'navigates to IULA
        End If

        EMWriteScreen "030", 12, 46

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

        TRANSMIT
        '----------------------------------------------------------------------------------------writing the note on IULB
    	EMReadScreen action_code_err_msg, 11, 24, 2
    	action_code_err_msg = trim(action_code_err_msg)
    	IF action_code_err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    	If action_code_err_msg = "ACTION CODE" THEN script_end_procedure_with_error_report(action_code_err_msg & vbNewLine & "Please ensure you are selecting the correct code for resolve. PF10 to ensure the match can be resolved using the script.")'checking for error msg'
    	Call clear_line_of_text(8, 6)
    	EMWriteScreen "Claim entered. See Case Note. ", 8, 6
    	Call clear_line_of_text(17, 9)
    	If action_header <> "ACTH" THEN EMWriteScreen Claim_number, 17, 9
    	'need to check about adding for multiple claims'
    	'msgbox "did the notes input?"
    	TRANSMIT 'this will take us back to IEVP main menu'
    	EMReadScreen err_msg, 75, 24, 2
    	err_msg = trim(err_msg)
    	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    	'ENTER COMMENTS TO EXPLAIN ACTION/S TAKEN
    	'PERSON IS NOT ASSOCIATED WITH THIS CLAIM '
    	'CLAIM MAY NOT EXIST FOR THE SELECTED ACTION - this is ACTH

        EMReadScreen panel_name, 4, 02, 52
        IF panel_name <> "IEVP" THEN 'msgbox "Script did not find IEVP."
            CALL back_to_SELF
            CALL navigate_to_MAXIS_screen("INFC" , "____")
            CALL write_value_and_transmit("IEVP", 20, 71)
            CALL write_value_and_transmit(SSN_number_read, 3, 63)
        End If
    '------------------------------------------------------------------back on the IEVP menu, making sure that the match cleared
        'msgbox panel_name
        EMReadScreen days_pending, 5, row, 72
        days_pending = trim(days_pending)
        match_cleared = TRUE
    	IF IsNumeric(days_pending) = TRUE THEN match_cleared = FALSE
    	If match_cleared = FALSE Then
        	confirm_cleared = MsgBox ("The script cannot identify that this match has cleared." & vbNewLine & vbNewLine & "Review IEVP and find the match that is being cleared with this run." &vbNewLine & " ** HAS THE MATCH BEEN CLEARED? **", vbQuestion + vbYesNo, "Confirm Match Cleared")
        	IF confirm_cleared = vbYes Then match_cleared = TRUE
    		IF confirm_cleared = vbno Then
    			match_cleared = FALSE
    			script_end_procedure_with_error_report("This match did not clear in IEVP, please advise what may have happened.")
    		END IF
    	End If

        IF match_type = "WAGE" THEN
        	IF select_quarter = 1 THEN IEVS_quarter = "1ST"
        	IF select_quarter = 2 THEN IEVS_quarter = "2ND"
         	IF select_quarter = 3 THEN IEVS_quarter = "3RD"
         	IF select_quarter = 4 THEN IEVS_quarter = "4TH"
        END IF
        IEVS_period = replace(IEVS_period, "/", " to ")
        Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days requested for HEADER of casenote'
        PF3 'back to the DAIL'
        PF3

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

    	IF ATR_needed_checkbox= CHECKED THEN header_note = "ATR/EVF STILL REQUIRED"
    	'------------------------------------------setting up case note header'
    	IF match_type = "BEER" THEN match_type_letter = "B"
    	IF match_type = "UBEN" THEN match_type_letter = "U"
    	IF match_type = "UNVI" THEN match_type_letter = "U"
    '-----------------------------------------------------------------------------------------CASENOTE
        start_a_blank_case_note
        IF match_type = "WAGE" THEN CALL write_variable_in_case_note("-----" & IEVS_quarter & " QTR " & IEVS_year & "WAGE MATCH"  & " (" & first_name & ") CLEARED CC-CLAIM ENTERED " & header_note & "-----")
        IF match_type = "BEER" THEN CALL write_variable_in_case_note("-----" & IEVS_year & " NON-WAGE MATCH(" & match_type_letter & ") " & " (" & first_name & ") CLEARED CC-CLAIM ENTERED " & header_note & "-----")
    	IF match_type = "UNVI" THEN CALL write_variable_in_case_note("-----" & IEVS_year & " NON-WAGE MATCH(" & match_type_letter & ") " & " (" & first_name & ") CLEARED CC-CLAIM ENTERED " & header_note & "-----")
        IF match_type = "UBEN" THEN CALL write_variable_in_case_note("-----" & IEVS_period & " NON-WAGE MATCH(" & match_type_letter & ") " & " (" & first_name & ") CLEARED CC-CLAIM ENTERED " & header_note & "-----")
    	CALL write_bullet_and_variable_in_case_note("Discovery date", discovery_date)
    	CALL write_bullet_and_variable_in_case_note("Period", IEVS_period)
    	CALL write_bullet_and_variable_in_case_note("Active Programs", programs)
    	CALL write_bullet_and_variable_in_case_note("Source of income", income_source)
    	CALL write_variable_in_case_note("----- ----- ----- ----- ----- ----- -----")
        Call write_variable_in_case_note(OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number & " Amt $" & Claim_amount)
        IF OP_program_II <> "Select:" then Call write_variable_in_case_note(OP_program_II & " Overpayment " & OP_from_II & " through " & OP_to_II & " Claim #" & Claim_number_II & " Amt $" & Claim_amount_II)
        IF OP_program_III <> "Select:" then Call write_variable_in_case_note(OP_program_III & " Overpayment " & OP_from_III & " through " & OP_to_III & " Claim #" & Claim_number_III & " Amt $" & Claim_amount_III)
        IF OP_program_IV <> "Select:" then Call write_variable_in_case_note(OP_program_IV & " Overpayment " & OP_from_IV & " through " & OP_to_IV & " Claim #" & Claim_number_IV & " Amt $" & Claim_amount_IV)
        IF HC_claim_number <> "" THEN
        	Call write_variable_in_case_note("HC OVERPAYMENT " & HC_from & " through " & HC_to & " Claim #" & HC_claim_number & " Amt $" & HC_Claim_amount)
        	Call write_bullet_and_variable_in_case_note("Health Care responsible members", HC_resp_memb)
        	Call write_bullet_and_variable_in_case_note("Total Federal Health Care amount", Fed_HC_AMT)
        	Call write_variable_in_case_note("* Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
        END IF
    	IF EI_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Earned Income Disregard Allowed")
        IF EI_checkbox = UNCHECKED THEN CALL write_variable_in_case_note("* Earned Income Disregard Not Allowed")
        CALL write_bullet_and_variable_in_case_note("Fraud referral made", fraud_referral)
        CALL write_bullet_and_variable_in_case_note("Income verification received", EVF_used)
        CALL write_bullet_and_variable_in_case_note("Date verification received", income_rcvd_date)
        CALL write_bullet_and_variable_in_case_note("Reason for overpayment", Reason_OP)
        CALL write_bullet_and_variable_in_case_note("Other responsible member(s)", OT_resp_memb)
        CALL write_variable_in_case_note("----- ----- ----- ----- ----- ----- -----")
        CALL write_variable_in_case_note("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1")
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
    '-----------------------------------------------------------------writing the CCOL case note'
    	msgbox "Navigating to CCOL to add case note, please contact the BlueZone Scripts team with any concerns."
    	Call navigate_to_MAXIS_screen("CCOL", "CLSM")
    	EMWriteScreen Claim_number, 4, 9
    	TRANSMIT
    	'NO CLAIMS WERE FOUND FOR THIS CASE, PROGRAM, AND STATUS
    	EMReadScreen error_check, 75, 24, 2	'making sure we can actually update this case.
    	error_check = trim(error_check)
    	If error_check <> "" then script_end_procedure_with_error_report(error_check & vbcr & "Unable to update this case. Please review case, and run the script again if applicable.")

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
        Call write_variable_in_CCOL_note_test(OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number & " Amt $" & Claim_amount)
        IF OP_program_II <> "Select:" then Call write_variable_in_CCOL_note_test(OP_program_II & " Overpayment " & OP_from_II & " through " & OP_to_II & " Claim #" & Claim_number_II & " Amt $" & Claim_amount_II)
        IF OP_program_III <> "Select:" then Call write_variable_in_CCOL_note_test(OP_program_III & " Overpayment " & OP_from_III & " through " & OP_to_III & " Claim #" & Claim_number_III & " Amt $" & Claim_amount_III)
        IF OP_program_IV <> "Select:" then Call write_variable_in_CCOL_note_test(OP_program_IV & " Overpayment " & OP_from_IV & " through " & OP_to_IV & " Claim #" & Claim_number_IV & " Amt $" & Claim_amount_IV)
        IF HC_claim_number <> "" THEN
        	Call write_variable_in_CCOL_note_test("HC OVERPAYMENT " & HC_from & " through " & HC_to & " Claim #" & HC_claim_number & " Amt $" & HC_Claim_amount)
        	Call write_bullet_and_variable_in_CCOL_NOTE_test("Health Care responsible members", HC_resp_memb)
        	Call write_bullet_and_variable_in_CCOL_NOTE_test("Total Federal Health Care amount", Fed_HC_AMT)
        	Call write_variable_in_CCOL_note_test("* Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
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
        script_end_procedure_with_error_report("Overpayment case note entered and copied to CCOL, please review the case to make sure the notes updated correctly." & vbcr & next_page)
ELSE'------------------------------------------------------------------------------------------end of overpayment script

    IF resolution_status = "CF-Future Save" THEN
	'-------------------------------------------------------------------------------------------------DIALOG
	    Dialog1 = "" 'Blanking out previous dialog detail
    	BeginDialog Dialog1, 0, 0, 161, 130, "Cleared CF Future Savings"
    	  DropListBox 65, 5, 90, 15, "Select One:"+chr(9)+"Case Became Ineligible"+chr(9)+"Person Removed"+chr(9)+"Benefit Increased"+chr(9)+"Benefit Decreased", IULB_result_dropdown
    	  DropListBox 65, 25, 90, 15, "One Time Only"+chr(9)+"Per Month For Nbr of Months", IULB_method_dropdown
    	  EditBox 65, 45, 40, 15, IULB_savings_amount
    	  EditBox 65, 70, 15, 15, IULB_start_month
    	  EditBox 90, 70, 15, 15, IULB_start_year
    	  EditBox 90, 90, 15, 15, IULB_months
    	  ButtonGroup ButtonPressed
    	    OkButton 50, 110, 50, 15
    	    CancelButton 105, 110, 50, 15
    	  Text 5, 10, 60, 10, "Results for IULB:"
    	  Text 5, 30, 55, 10, "Method for IULB:"
    	  Text 5, 50, 55, 10, "Savings Amount:"
    	  Text 65, 60, 15, 10, "MM"
    	  Text 90, 60, 10, 10, "YY"
    	  Text 5, 75, 35, 10, "Start Date:"
    	  Text 5, 95, 70, 10, "Months for Method R:"
    	EndDialog

	    DO
	    	err_msg = ""
	    	Dialog Dialog1
	    	cancel_without_confirmation
	    	IF IsNumeric(IULB_savings_amount) = false THEN err_msg = err_msg & vbNewLine & "* Please enter a valid numeric amount no decimal."
	    	IF IULB_result = "Select One:" THEN err_msg = err_msg & vbNewLine & "Please enter the IULB result."
	    	IF IULB_method = "Select One:" THEN err_msg = err_msg & vbNewLine & "Please enter the IULB method."
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
'----------------------------------------------------------------------------------------------------RESOLVING THE MATCH
    'msgbox "are we writing?"
	EMWriteScreen resolve_time, 12, 46	    'resolved notes depending on the resolution_status
	IF resolution_status = "CB-Ovrpmt And Future Save" THEN IULA_res_status = "CB"
	'IF resolution_status = "CC-Overpayment Only" THEN IULA_res_status = "CC" 'Claim Entered" CC cannot be used - ACTION CODE FOR ACTH OR ACTM IS INVALID
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
	'msgbox IULA_res_status
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
	IF resolution_status = "CB-Ovrpmt And Future Save" THEN EMWriteScreen "OP Claim entered and future savings." & other_notes, 8, 6
	IF resolution_status = "CC-Overpayment Only" THEN EMWriteScreen "OP Claim entered." & other_notes, 8, 6
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

	'msgbox "did the notes input?"
	TRANSMIT 'this will take us back to IEVP main menu'
'--------------------------------------------------------------------back on the IEVP menu, making sure that the match cleared
	EMReadScreen days_pending, 5, 7, 72
	days_pending = trim(days_pending)
	IF IsNumeric(days_pending) = TRUE THEN
		match_cleared = FALSE
		script_end_procedure_with_error_report("This match did not appear to clear. Please check case, and try again.")
	ELSE
	  match_cleared = TRUE
	END IF
	pending_verifs = trim(pending_verifs) 	'takes the last comma off of pending_verifs when autofilled into dialog if more more than one app date is found and additional app is selected
	IF right(pending_verifs, 1) = "," THEN pending_verifs = left(pending_verifs, len(pending_verifs) - 1)
	IF match_type = "WAGE" THEN
		IF select_quarter = 1 THEN IEVS_quarter = "1ST"
		IF select_quarter = 2 THEN IEVS_quarter = "2ND"
		IF select_quarter = 3 THEN IEVS_quarter = "3RD"
		IF select_quarter = 4 THEN IEVS_quarter = "4TH"
	END IF
'------------------------------------------------------------------setting up case note header'
	IEVS_period = trim(IEVS_period)
	IF match_type <> "UBEN" THEN IEVS_period = replace(IEVS_period, "/", " to ")
	IF match_type = "UBEN" THEN IEVS_period = replace(IEVS_period, "-", "/")
	Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days

'------------------------------------------------------------------STAT/MISC for claim referral tracking
	IF no_overpayment_checkbox = CHECKED or overpayment_exists_checkbox = CHECKED THEN
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
		IF overpayment_exists_checkbox = CHECKED THEN MISC_action_taken =  "Determination-OP Entered" '"Claim Determination 25 character available
		IF no_overpayment_checkbox = CHECKED THEN MISC_action_taken = "Determination-No Savings"
		EMWriteScreen MISC_action_taken, Row, 30
		EMWriteScreen date, Row, 66
        TRANSMIT
        '-------------------------------------------------------------------------------------------------The case note
        start_a_blank_case_note
        Call write_variable_in_case_note("-----Claim Referral Tracking - Initial Claim Referral-----")
	    Call write_bullet_and_variable_in_case_note("Action Date", action_date)
        Call write_bullet_and_variable_in_case_note("Active Program(s)", programs)
        IF next_action = "Sent Request for Additional Info" THEN CALL write_variable_in_case_note("* Additional verifications requested, follow up set for 10 day return.")
        Call write_bullet_and_variable_in_case_note("Other Notes", other_notes)
        Call write_variable_in_case_note("* Entries for these potential claims must be retained until further notice.")
	    IF case_note_only = TRUE THEN Call write_variable_in_case_note("Maxis case is inactive unable to add or update MISC panel")
        Call write_variable_in_case_note("-----")
        Call write_variable_in_case_note(worker_signature)
        PF3
	END IF

   '----------------------------------------------------------------the case match CLEARED note
	start_a_blank_case_note
	IF resolution_status <> "NC-Non Cooperation" THEN
		IF match_type = "WAGE" THEN CALL write_variable_in_case_note("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH (" & first_name & ") CLEARED " & IULA_res_status & "-----")
		IF match_type = "BEER" THEN CALL write_variable_in_case_note("-----" & IEVS_year & " NON-WAGE MATCH (" & first_name & ") (B) CLEARED " & IULA_res_status & "-----")
		IF match_type = "UNVI" THEN CALL write_variable_in_case_note("-----" & IEVS_year & " NON-WAGE MATCH (" & first_name & ") (U) CLEARED " & IULA_res_status & "-----")
		IF match_type = "UBEN" THEN CALL write_variable_in_case_note("-----" & IEVS_period & " NON-WAGE MATCH (" & first_name & ") (U) CLEARED " & IULA_res_status & "-----")
		CALL write_bullet_and_variable_in_case_note("Period", IEVS_period)
		CALL write_bullet_and_variable_in_case_note("Active Programs", programs)
		CALL write_bullet_and_variable_in_case_note("Source of income", income_source)
		CALL write_variable_in_case_note ("----- ----- -----")
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
  		IF change_response <> "N/A" THEN CALL write_bullet_and_variable_in_case_note("Responded to Difference Notice", change_response)
  		CALL write_bullet_and_variable_in_case_note("Other notes", other_notes)
  		CALL write_variable_in_case_note("----- ----- ----- ----- -----")
  		CALL write_variable_in_case_note ("DEBT ESTABLISHMENT UNIT 612-348-4290 EXT 1-1-1")
		PF3
	ELSEIF resolution_status = "NC-Non Cooperation" THEN   '
		IF match_type = "WAGE" THEN CALL write_variable_in_case_note("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH (" & first_name & ") NON-COOPERATION-----")
		IF match_type = "BEER" THEN CALL write_variable_in_case_note("-----" & IEVS_year & " NON-WAGE MATCH (" & first_name & ") (B) NON-COOPERATION-----")
		IF match_type = "UNVI" THEN CALL write_variable_in_case_note("-----" & IEVS_year & " NON-WAGE MATCH (" & first_name & ") (U) NON-COOPERATION-----")
		IF match_type = "UBEN" THEN CALL write_variable_in_case_note("-----" & IEVS_period & " NON-WAGE MATCH (" & first_name & ") (U) NON-COOPERATION-----")
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
		CALL write_bullet_and_variable_in_case_note("Other notes", other_notes)
		CALL write_variable_in_case_note("----- ----- ----- ----- -----")
		CALL write_variable_in_case_note ("DEBT ESTABLISHMENT UNIT 612-348-4290 EXT 1-1-1")
		PF3
	END IF
	'-------------------------------The following will generate a TIKL formatted date for 10 days from now, and add it to the TIKL
	IF tenday_checkbox = 1 THEN Call create_TIKL("Unable to close due to 10 day cutoff. Verification of match should have returned by now. If not received and processed, take appropriate action.", 0, date, True)
	script_end_procedure_with_error_report("Match has been acted on. Please take any additional action needed for your case.")
END IF
