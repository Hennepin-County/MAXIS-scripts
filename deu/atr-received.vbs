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
CALL changelog_update("10/06/2022", "Update to remove hard coded DEU signature all DEU scripts.", "MiKayla Handley, Hennepin County") '#316
CALL changelog_update("06/21/2022", "Updated handling for non-disclosure agreement and closing documentation.", "MiKayla Handley, Hennepin County") '#493
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

'---------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 201, 65, "ATR RECEIVED"
  EditBox 70, 5, 50, 15, MAXIS_case_number
  EditBox 70, 25, 125, 15, worker_signature
  Text 5, 30, 60, 10, "Worker Signature:"
  ButtonGroup ButtonPressed
    OkButton 100, 45, 45, 15
    CancelButton 150, 45, 45, 15
  Text 5, 10, 50, 10, "Case Number:"
EndDialog

DO
	DO
		err_msg = ""
		DIALOG Dialog1
		cancel_without_confirmation
  		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		IF trim(worker_signature) = "" THEN err_msg = err_msg & vbNewLine & "Please enter your worker signature."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)
LOOP UNTIL are_we_passworded_out = false

'----------------------------------------------------------------------------------------------------Gathering the member information
CALL navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_priv)
IF is_this_priv = TRUE THEN script_end_procedure("This case is privileged, the script will now end.")
client_array = "Select One:" & "|"

DO	'reads the reference number, last name, first name, and THEN puts it into a single string THEN into the array
	EMReadscreen ref_nbr, 3, 4, 33
	EMReadScreen access_denied_check, 13, 24, 2
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
CALL navigate_to_MAXIS_screen("INFC" , "____")
CALL write_value_and_transmit("IEVP", 20, 71)
CALL write_value_and_transmit(IEVS_ssn, 3, 63)
'checking for NON-DISCLOSURE AGREEMENT REQUIRED FOR ACCESS TO IEVS FUNCTIONS'
EMReadScreen agreement_check, 9, 2, 24
IF agreement_check = "Automated" THEN script_end_procedure("To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the screens and review the agreement.")

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
    EMReadScreen number_IEVS_type, 3, 7, 12 'read the DAIL msg'
    IF number_IEVS_type = "A30" THEN match_type = "BNDX"
    'IF number_IEVS_type = "A40" THEN match_type = "SDXS/I"
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
		select_quarter = "YEAR"
	END IF
END IF

EMReadScreen number_IEVS_type, 3, 7, 12 'read the DAIL msg'
IF number_IEVS_type = "A30" THEN match_type = "BNDX"
IF number_IEVS_type = "A40" THEN match_type = "SDXS/I"
IF number_IEVS_type = "A70" THEN match_type = "BEER"
IF number_IEVS_type = "A80" THEN match_type = "UNVI"
IF number_IEVS_type = "A60" THEN match_type = "UBEN"
IF number_IEVS_type = "A50" or number_IEVS_type = "A51"  THEN match_type = "WAGE"

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

ATR_Verf_CheckBox = CHECKED

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 271, 240, "ATR RECEIVED FOR: "  & MAXIS_case_number
  CheckBox 175, 15, 70, 10, "Difference Notice", Diff_Notice_Checkbox
  CheckBox 175, 25, 90, 10, "Authorization to Release", ATR_Verf_CheckBox
  CheckBox 175, 45, 90, 10, "Employment Verification", EVF_checkbox
  CheckBox 175, 55, 80, 10, "Lottery/Gaming Form", lottery_verf_checkbox
  CheckBox 175, 65, 80, 10, "Rental Income Form", rental_checkbox
  CheckBox 175, 35, 85, 10, "First Page of EVF (only)", pageone_EVF_checkbox
  CheckBox 175, 75, 80, 10, "Other (please specify)", other_checkbox
  DropListBox 85, 105, 60, 15, "Select One:"+chr(9)+"Deleted DISQ"+chr(9)+"Pending verif"+chr(9)+"N/A", DISQ_action
  EditBox 210, 90, 45, 15, date_ATR_received
  EditBox 55, 125, 95, 15, source_address
  EditBox 55, 145, 45, 15, source_phone
  CheckBox 150, 150, 115, 10, "Set a TIKL due to 10 day cutoff", tenday_checkbox
  DropListBox 140, 175, 115, 15, "Select One:"+chr(9)+"Not Needed"+chr(9)+"Initial"+chr(9)+"Overpayment Exists"+chr(9)+"OP Non-Collectible (please specify)"+chr(9)+"No Savings/Overpayment", claim_referral_tracking_dropdown
  EditBox 50, 200, 215, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 175, 220, 45, 15
    CancelButton 220, 220, 45, 15
  Text 5, 5, 165, 10, "Client Name: "   & client_name
  Text 5, 20, 150, 10, "Match Type: "  & match_type
  Text 5, 35, 150, 10, "Match Period: "  & IEVS_period
  Text 5, 50, 160, 10, "Active Programs: "   & programs
  Text 5, 65, 165, 20, "Income Source: "     & income_source
  Text 5, 90, 120, 10, "Difference Notice: "   & sent_date
  GroupBox 170, 5, 95, 105, "Verification(s) Received: "
  Text 5, 110, 75, 10, "DISQ panel addressed:"
  Text 185, 95, 20, 10, "Date:"
  Text 5, 130, 30, 10, "Address:"
  Text 5, 150, 45, 10, "Fax or Phone:"
  GroupBox 5, 165, 260, 30, "SNAP or MFIP Federal Food only"
  Text 10, 180, 130, 10, "Claim Referral Tracking on STAT/MISC:"
  Text 5, 205, 40, 10, "Other notes: "
EndDialog

DO
    DO
        err_msg = ""
    	Dialog Dialog1
    	cancel_without_confirmation
    	IF date_ATR_received = "" THEN err_msg = err_msg & vbNewLine & "Please provide the date the ATR was received"
    	IF DISQ_action = "Select One:" THEN err_msg = err_msg & vbNewLine & "Please advise if DISQ panel was updated"
		IF claim_referral_tracking_dropdown =  "Select One:" THEN err_msg = err_msg & vbNewLine & "Please select if the claim referral tracking needs to be updated."
    	IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    LOOP UNTIL err_msg = ""
    CALL check_for_password_without_transmit(are_we_passworded_out)
LOOP UNTIL are_we_passworded_out = false


EMwritescreen "005", 12, 46 'writing the resolve time to read for later
IF notice_sent = "Y" THEN
	EMwritescreen "Y", 15, 37 'Responded to diff notice
ELSE
	EMwritescreen "N", 15, 37 'Responded to diff notice
END IF

TRANSMIT 'this will take us to IULB'

ROW = 8
EMReadScreen IULB_first_line, 1, row, 6
IF IULB_first_line = "" THEN
	EMwritescreen "ATR RECEIVED " & date_ATR_received, row, 6
ELSE
	ROW = 9
	CALL clear_line_of_text(row, 6)
	EMwritescreen "ATR RECEIVED " & date_ATR_received, row, 6
END IF

TRANSMIT 'exiting IULA, helps prevent errors when going to the case note

pending_verifs = ""
IF Diff_Notice_Checkbox = CHECKED THEN pending_verifs = pending_verifs & "Difference Notice, "
IF empl_verf_checkbox = CHECKED THEN pending_verifs = pending_verifs & "EVF, "
IF ATR_Verf_CheckBox = CHECKED THEN pending_verifs = pending_verifs & "ATR, "
IF lottery_verf_checkbox = CHECKED THEN pending_verifs = pending_verifs & "Lottery/Gaming Form, "
IF rental_checkbox =  CHECKED THEN pending_verifs = pending_verifs & "Rental Income Form, "
IF other_checkbox = CHECKED THEN pending_verifs = pending_verifs & "Other, "

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

IF match_type = "BEER" THEN match_type_letter = "B"
IF match_type = "UBEN" THEN match_type_letter = "U"
IF match_type = "UNVI" THEN match_type_letter = "U"

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
END IF
'-------------------------------------------------------------------------------------------------The case note
start_a_blank_case_note
IF match_type = "WAGE" THEN CALL write_variable_in_case_note("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH"  & " (" & first_name & ") " & "ATR RECEIVED-----")
IF match_type = "BEER" THEN CALL write_variable_in_case_note("-----" & IEVS_year & " NON-WAGE MATCH(" & match_type_letter & ") " & " (" & first_name & ") " & "ATR RECEIVED-----")
IF match_type = "UNVI" THEN CALL write_variable_in_case_note("-----" & IEVS_year & " NON-WAGE MATCH(" & match_type_letter & ") " & " (" & first_name & ") " & "ATR RECEIVED-----")
IF match_type = "UBEN" THEN CALL write_variable_in_case_note("-----" & IEVS_period & " NON-WAGE MATCH(" & match_type_letter & ") " & " (" & first_name & ") " & "ATR RECEIVED-----")
CALL write_variable_in_CASE_NOTE("* Date ATR received: " & date_ATR_received)
CALL write_bullet_and_variable_in_case_note("Discovery date", discovery_date)
CALL write_bullet_and_variable_in_case_note("Period", IEVS_period)
CALL write_bullet_and_variable_in_case_note("Active Programs", programs)
CALL write_bullet_and_variable_in_case_note("Source of income", income_source)
CALL write_variable_in_CASE_NOTE ("----- ----- ----- ----- -----")
IF DISQ_action = "Pending verif" THEN CALL write_variable_in_CASE_NOTE("* Pending verification of income or asset.")
IF DISQ_action = "Deleted DISQ" THEN CALL write_variable_in_CASE_NOTE("* Updated DISQ panel.")
CALL write_bullet_and_variable_in_case_note("Verifications Received", pending_verifs)
CALL write_bullet_and_variable_in_case_note("Source Address:", source_address)
CALL write_bullet_and_variable_in_case_note("Fax/Phone:", source_phone)
CALL write_bullet_and_variable_in_case_note("Response to Difference Notice", notice_sent)
IF notice_sent = "Y" THEN CALL write_variable_in_CASE_NOTE("* IEVP updated as responded to difference notice")
CALL write_bullet_and_variable_in_case_note("Other Notes", other_notes)
IF DISQ_action <> "Pending verif" THEN CALL write_variable_in_CASE_NOTE("---The case may be eligible for REIN if all necessary paperwork has been received.")
CALL write_variable_in_CASE_NOTE ("----- ----- ----- ----- -----")
CALL write_variable_in_case_note(worker_signature)

script_end_procedure_with_error_report("ATR case note updated successfully." & vbNewLine & "Please remember to update/delete the DISQ panel")
'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------06/24/2022
'--Tab orders reviewed & confirmed----------------------------------------------06/24/2022
'--Mandatory fields all present & Reviewed--------------------------------------06/24/2022
'--All variables in dialog match mandatory fields-------------------------------06/24/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------06/24/2022
'--CASE:NOTE Header doesn't look funky------------------------------------------06/24/2022
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------06/24/2022
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------06/24/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------06/24/2022
'--PRIV Case handling reviewed -------------------------------------------------06/24/2022
'--Out-of-County handling reviewed----------------------------------------------06/24/2022------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------06/24/2022
'--BULK - review output of statistics and run time/count (if applicable)--------06/24/2022------------------N/A
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---06/24/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------06/24/2022------------------N/A
'--Incrementors reviewed (if necessary)-----------------------------------------06/24/2022
'--Denomination reviewed -------------------------------------------------------06/24/2022
'--Script name reviewed---------------------------------------------------------06/24/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------06/24/2022
'--Comment Code-----------------------------------------------------------------06/24/2022
'--Update Changelog for release/update------------------------------------------06/24/2022
'--Remove testing message boxes-------------------------------------------------06/24/2022
'--Remove testing code/unnecessary code-----------------------------------------06/24/2022
'--Review/update SharePoint instructions----------------------------------------06/24/2022 
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------06/24/2022
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------06/24/2022
'--Complete misc. documentation (if applicable)---------------------------------06/24/2022
'--Update project team/issue contact (if applicable)----------------------------06/24/2022
