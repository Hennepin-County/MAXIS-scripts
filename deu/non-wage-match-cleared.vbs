''GATHERING STATS===========================================================================================
name_of_script = "ACTION - DEU NONWAGE MATCH CLEARED.vbs"
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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
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
CALL changelog_update("05/14/2018", "Updated to add CF drop down and handling for UBEN matches.", "MiKayla Handley, Hennepin County")
CALL changelog_update("05/08/2018", "Updated to pull employer name correctly.", "MiKayla Handley, Hennepin County")
CALL changelog_update("02/13/2018", "Updated for clearing UBEN CC.", "MiKayla Handley, Hennepin County")
CALL changelog_update("01/16/2018", "Corrected casenote for pulling IEVS period.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/27/2017", "Updated to handle clearing the match when the date is over 45 days.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/27/2017", "Updated to handle clearing the match BE-OP entered.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/14/2017", "Initial version.", "MiKayla Handley, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

function DEU_password_check(end_script)
'--- This function checks to ensure the user is in a MAXIS panel
'~~~~~ end_script: If end_script = TRUE the script will end. If end_script = FALSE, the user will be given the option to cancel the script, or manually navigate to a MAXIS screen.
'===== Keywords: MAXIS, production, script_end_procedure
	Do
		EMReadScreen MAXIS_check, 5, 1, 39
		If MAXIS_check <> "MAXIS"  and MAXIS_check <> "AXIS " then
			If end_script = True then
				script_end_procedure("You do not appear to be in MAXIS. You may be passworded out. Please check your MAXIS screen and try again.")
			Else
				warning_box = MsgBox("You do not appear to be in MAXIS. You may be passworded out. Please check your MAXIS screen and try again, or press ""cancel"" to exit the script.", vbOKCancel)
				If warning_box = vbCancel then stopscript
			End if
		End if
	Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "
end function

'---------------------------------------------------------------------THE SCRIPT
EMConnect ""
'--------------------------------------------------------------------CHECKS TO MAKE SURE THE WORKER IS ON THEIR DAIL
EMReadscreen dail_check, 4, 2, 48
IF dail_check <> "DAIL" THEN script_end_procedure("You are not in your dail. This script will stop.")
EMSendKey "t"
'checking for an active MAXIS session
Call check_for_MAXIS(FALSE)
EMReadScreen IEVS_type, 4, 6, 6 'read the DAIL msg'
'msgbox IEVS_type
IF IEVS_type = "UBEN" or IEVS_type = "BEER" THEN
	match_found = TRUE
ELSE
	script_end_procedure("This is not a IEVS match. Please select a WAGE match DAIL, and run the script again.")
END IF
IF IEVS_type = "BEER" THEN type_match = "B"
IF IEVS_type = "UBEN" THEN type_match = "U"

EMReadScreen MAXIS_case_number, 8, 5, 73
MAXIS_case_number= TRIM(MAXIS_case_number)

'----------------------------------------------------------------------------------------------------IEVS
'Navigating deeper into the match interface
CALL write_value_and_transmit("I", 6, 3)   		'navigates to INFC
CALL write_value_and_transmit("IEVP", 20, 71)   'navigates to IEVP
IF error_msg = "NO IEVS" THEN script_end_procedure("An error occurred in IEVP, please process manually.")'checking for error msg'
EMReadScreen IEVS_match, 11, row, 47

'-------------------------------------------------------------------Ensuring that match has not already been resolved.
Row = 7
DO
	EMReadScreen IEVS_match, 11, row, 47
	IF trim(IEVS_match) = "" THEN script_end_procedure("IEVS match for the selected period could not be found. The script will now end.")
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
	IF ievp_info_confirmation = vbCancel THEN script_end_procedure ("The script has ended. The match has not been acted on.")
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

'---------------------------------------------------------------------Reading potential errors for out-of-county cases
EMReadScreen OutOfCounty_error, 12, 24, 2
IF OutOfCounty_error = "MATCH IS NOT" THEN
	script_end_procedure("Out-of-county case. Cannot update.")
ELSE
	IF IEVS_type = "UBEN" THEN
		EMReadScreen UBEN_month, 2, 5, 68
		EMReadScreen UBEN_year, 2, 5, 71
		'EMReadScreen source_income, 29, 8, 37
		source_income = "Unemployment"
		'source_income = trim(source_income)
		'IF instr(source_income, " AMT: $") THEN 					  'establishing the length of the variable
		    'position = InStr(source_income, " AMT: $")    		      'sets the position at the deliminator
 				'source_income = Left(source_income, position)
		'END IF
	ELSEIF IEVS_type = "BEER" THEN
		EMReadScreen IEVS_year, 2, 8, 15
		IEVS_year = "20" & IEVS_year
		EMReadScreen source_income, 42, 8, 28
		source_income = trim(source_income)
		IF instr(source_income, " AMOUNT: $") THEN
		    position = InStr(source_income, " AMOUNT: $")    		      'sets the position at the deliminator
		    source_income = Left(source_income, position)
		END IF
	END IF
END IF

'--------------------------------------------------------------------Client name
EMReadScreen client_name, 35, 5, 24
client_name = trim(client_name)                         'trimming the client name
IF instr(client_name, ",") THEN    						'Most cases have both last name and 1st name. This seperates the two names
	length = len(client_name)                           'establishing the length of the variable
	position = InStr(client_name, ",")                  'sets the position at the deliminator (in this case the comma)
	last_name = Left(client_name, position-1)           'establishes client last name as being before the deliminator
	first_name = Right(client_name, length-position)    'establishes client first name as after before the deliminator
ELSE                                'In cases where the last name takes up the entire space, THEN the client name becomes the last name
	first_name = ""
	last_name = client_name

END IF
IF instr(first_name, " ") THEN   						'If there is a middle initial in the first name, THEN it removes it
	length = len(first_name)                        	'trimming the 1st name
	position = InStr(first_name, " ")               	'establishing the length of the variable
	first_name = Left(first_name, position-1)       	'trims the middle initial off of the first name
END IF
'it is not putting a space in'
'----------------------------------------------------------------------------------------------------ACTIVE PROGRAMS
EMReadScreen Active_Programs, 13, 6, 68
Active_Programs =trim(Active_Programs)

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
EMReadScreen notice_sent, 1, 14, 37
EMReadScreen sent_date, 8, 14, 68
sent_date = trim(sent_date)
IF sent_date = "" THEN sent_date = replace(sent_date, " ", "N/A")
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
  Text 120, 40, 165, 15, "Income source: "   & source_income
  Text 5, 125, 40, 10, "Other notes: "
EndDialog

IF notice_sent = "N" THEN
	DO
	 	  err_msg = ""
      Dialog notice_action_dialog
      cancel_confirmation
      IF (send_notice_checkbox = UNCHECKED AND clear_action_checkbox = UNCHECKED) THEN err_msg = err_msg & vbNewLine & "* Please select an answer to continue."
      IF (send_notice_checkbox = CHECKED AND clear_action_checkbox = CHECKED) THEN err_msg = err_msg & vbNewLine & "* Please select only one answer to continue."
      IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
END IF

CALL DEU_password_check(False)

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
	CALL DEU_password_check(False)

	'--------------------------------------------------------------------sending the notice in IULA
	EMwritescreen "005", 12, 46 'writing the resolve time to read for later
	EMwritescreen "Y", 14, 37 'send Notice
	msgbox "Difference Notice Sent"
	transmit 'goes into IULA
	ROW = 8
    EMReadScreen IULB_first_line, 1, row, 6
    IF IULB_first_line = "" THEN
    	EMwritescreen "Difference Notice Sent", row, 6
    ELSE
    	ROW = 9
    	CALL clear_line_of_text(row, 6)
    	EMwritescreen "Difference Notice Sent", row, 6
    END IF
	transmit'exiting IULA, helps prevent errors when going to the case note
	'--------------------------------------------------------------------The case note & case note related code
	pending_verifs = ""
  	IF Diff_Notice_Checkbox = CHECKED THEN pending_verifs = pending_verifs & "Difference Notice, "
	IF empl_verf_checkbox = CHECKED THEN pending_verifs = pending_verifs & "EVF, "
	IF ATR_Verf_CheckBox = CHECKED THEN pending_verifs = pending_verifs & "ATR, "
	IF other_checkbox = CHECKED THEN pending_verifs = pending_verifs & "Other, "

    '-------------------------------------------------------------------trims excess spaces of pending_verifs
  	pending_verifs = trim(pending_verifs) 	'takes the last comma off of pending_verifs when autofilled into dialog if more more than one app date is found and additional app is selected
  	IF right(pending_verifs, 1) = "," THEN pending_verifs = left(pending_verifs, len(pending_verifs) - 1)

	IEVS_match = replace(IEVS_match, "/", " to ")
	Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days

	'---------------------------------------------------------------------DIFF NOTC case note
  	start_a_blank_CASE_NOTE
		IF IEVS_type = "UBEN" THEN CALL write_variable_in_CASE_NOTE("-----" & UBEN_month & "/" & UBEN_year & " NON-WAGE MATCH (" & type_match & ") " & "(" & first_name & ") DIFF NOTICE SENT-----")
		IF IEVS_type = "BEER" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_year & " NON-WAGE MATCH (" & type_match & ") " & "(" & first_name & ") DIFF NOTICE SENT-----")
		CALL write_bullet_and_variable_in_CASE_NOTE("Client Name", Client_Name)
  		CALL write_bullet_and_variable_in_CASE_NOTE("Active Programs", programs)
		CALL write_bullet_and_variable_in_CASE_NOTE("Source of income", source_income)
		CALL write_variable_in_CASE_NOTE ("----- ----- -----")
  		CALL write_bullet_and_variable_in_CASE_NOTE("Verification Requested", pending_verifs)
  		CALL write_bullet_and_variable_in_CASE_NOTE("Verification Due", Due_date)
		CALL write_variable_in_CASE_NOTE ("* Client must be provided 10 days to return requested verifications *")
  		CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
  		CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- ----- ----- -----")
  		CALL write_variable_in_CASE_NOTE ("DEBT ESTABLISHMENT UNIT 612-348-4290 EXT 1-1-1")
END IF

IF clear_action_checkbox = CHECKED or notice_sent = "Y" THEN
	'IF sent_date = "" THEN replace(sent_date, "N/A")
	'MsgBox("A difference notice was sent: " & sent_date & "." & vbNewLine & "The script will now navigate to clear the Non-wage match.")
  	BeginDialog cleared_match_dialog, 0, 0, 311, 245, "NON-WAGE MATCH CLEARED"
    	Text 10, 20, 110, 10, "Case number: " & MAXIS_case_number
    	Text 120, 20, 165, 10, "Client name: " & client_name
    	Text 10, 40, 105, 10, "Active Programs: " & programs
    	Text 120, 40, 175, 15, "Income source: " & source_income
    	DropListBox 75, 65, 110, 15, "Select One:"+chr(9)+"BC - Case Closed"+chr(9)+"BN - Already known, No Savings"+chr(9)+"BE - Child"+chr(9)+"BE - No Change"+chr(9)+"BE - OP entered"+chr(9)+"BO - Other"+chr(9)+"BP - Wrong Person"+chr(9)+"CC - Claim Entered"+chr(9)+"CF - Future Savings"+chr(9)+"NC - Non Cooperation", resolution_status
    	DropListBox 125, 85, 60, 15, "Select One:"+chr(9)+"Yes"+chr(9)+"No", change_response
    	EditBox 125, 105, 35, 15, resolve_time
    	CheckBox 210, 75, 70, 10, "Difference Notice", Diff_Notice_Checkbox
    	CheckBox 210, 85, 90, 10, "Authorization to Release", ATR_Verf_CheckBox
    	CheckBox 210, 95, 90, 10, "Employment verification", EVF_checkbox
    	CheckBox 210, 105, 80, 10, "Other (please specify)", other_checkbox
    	DropListBox 40, 140, 40, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"HC"+chr(9)+"MF", claim_program
    	EditBox 105, 140, 15, 15, from_month
    	EditBox 125, 140, 15, 15, from_year
    	EditBox 155, 140, 15, 15, to_month
    	EditBox 175, 140, 15, 15, to_year
    	EditBox 200, 140, 45, 15, claim_number
    	EditBox 255, 140, 45, 15, claim_AMT
    	DropListBox 50, 160, 35, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", collectible_dropdown
    	EditBox 155, 160, 145, 15, collectible_reason
    	EditBox 65, 180, 60, 15, Discovery_date
    	EditBox 215, 180, 85, 15, reason_OP
    	EditBox 55, 205, 245, 15, other_notes
    	CheckBox 10, 225, 135, 10, "Check here if 10 day cutoff has passed", TIKL_checkbox
    	ButtonGroup ButtonPressed
      		OkButton 205, 225, 45, 15
      		CancelButton 255, 225, 45, 15
    	GroupBox 5, 5, 300, 55, "NON-WAGE MATCH"
    	Text 10, 70, 60, 10, "Resolution Status: "
    	GroupBox 195, 65, 110, 55, "Verification Used to Clear: "
    	Text 10, 90, 110, 10, "Responded to Difference Notice: "
    	Text 10, 110, 85, 10, "Resolve time (in minutes): "
    	GroupBox 5, 125, 300, 75, "If resolution status is CC"
    	Text 10, 145, 30, 10, "Program: "
    	Text 85, 145, 20, 10, "from: "
    	Text 105, 130, 20, 10, "(MM) "
    	Text 125, 130, 15, 10, "(YY) "
    	Text 155, 130, 20, 10, "(MM) "
    	Text 175, 130, 15, 10, "(YY) "
    	Text 200, 130, 30, 10, "Claim #: "
    	Text 255, 130, 20, 10, "AMT: "
    	Text 145, 145, 10, 10, "to: "
    	Text 10, 165, 40, 10, "Collectible: "
    	Text 90, 165, 65, 10, "Reason collectible: "
    	Text 10, 185, 55, 10, "Discovery date: "
    	Text 130, 185, 85, 10, "Reason for overpayment:"
    	Text 10, 210, 40, 10, "Other notes: "
  	EndDialog

	DO
		err_msg = ""
		Dialog cleared_match_dialog
		cancel_confirmation
		IF IsNumeric(resolve_time) = false or len(resolve_time) > 3 THEN err_msg = err_msg & vbNewLine & "* Enter a valid numeric resolved time."
		IF resolve_time = "" THEN err_msg = err_msg & vbNewLine & "Please complete resolve time."
		IF change_response = "Select One:" THEN err_msg = err_msg & vbNewLine & "Did the client respond to Difference Notice?"
		IF resolution_status = "Select One:" THEN err_msg = err_msg & vbNewLine & "Please select a resolution status to continue."
		IF (resolution_status = "BE - No Change" AND other_notes = "") THEN err_msg = err_msg & vbNewLine & "When clearing using BE other notes must be completed."
		IF (resolution_status = "CC - Claim Entered" AND Reason_OP = "") THEN err_msg = err_msg & vbnewline & "* You must enter the reason for the overpayment."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL DEU_password_check(False)

	IF IEVS_type = "UBEN" THEN source_income = replace(source_income, "", "UNEA AMOUNTS NOT EQUAL")
	'----------------------------------------------------------------------------------------------------RESOLVING THE MATCH
	EMWriteScreen resolve_time, 12, 46

	'resolved notes depending on the resolution_status
	IF resolution_status = "BC - Case Closed" THEN rez_status = "BC"
	IF resolution_status = "BE - Child" THEN rez_status = "BE"
	IF resolution_status = "BE - No Change" THEN rez_status = "BE"
	IF resolution_status = "BE - OP Entered" THEN rez_status = "BE"
	IF resolution_status = "BN - Already known, No Savings" THEN rez_status = "BN"
	IF resolution_status = "BO - Other" THEN rez_status = "BO"
	IF resolution_status = "BP - Wrong Person"  THEN rez_status = "BP"
	IF resolution_status = "CC - Claim Entered" THEN rez_status = "CC"
	IF resolution_status = "CF - Future Savings" THEN rez_status = "CF"
	IF resolution_status = "NC - Non Cooperation" THEN rez_status = "NC"
	'CC cannot be used - ACTION CODE FOR ACTH OR ACTM IS INVALID
	programs_array = split(programs, ",")
	For each program in programs_array
		program = trim(program)
		IF program = "DWP" then cleared_header = "ACTD"
		IF program = "Food Support" then cleared_header = "ACTF"
		IF program = "Health Care" then cleared_header = "ACTH"
		IF program = "Medical Assistance" then cleared_header = "ACTM"
		IF program = "MFIP" then cleared_header = "ACTS"
		row = 11
		col = 57
		EMSearch cleared_header, row, col
		EMWriteScreen resolution_status, row + 1, col + 1
	Next

	EMwritescreen rez_status, 12, 58
	IF change_response = "YES" THEN
		EMwritescreen "Y", 15, 37
	ELSE
		EMwritescreen "N", 15, 37
	END IF
	transmit 'IULB
	IF resolution_status = "CF - Future Savings" THEN
  		BeginDialog future_savings_dialog, 0, 0, 161, 105, "Future Savings"
    		DropListBox 35, 5, 120, 15, "Select One:"+chr(9)+"I - Case Became Ineligible"+chr(9)+"R - Person Removed"+chr(9)+"P - Benefit Increased"+chr(9)+"N - Benefit Decreased"+chr(9)+"M - AFDC Closed w/ Extended Med"+chr(9)+"X - MFIP closed w/ Extended Med", saving_result
    		DropListBox 35, 25, 120, 15, "Select One:"+chr(9)+"O - One Time Only"+chr(9)+"R - Per Month For Nbr of Months", saving_method
			EditBox 35, 45, 40, 15, saving_amount
  			EditBox 120, 45, 15, 15, date_month
  			EditBox 140, 45, 15, 15, date_year
  			EditBox 140, 65, 15, 15, saving_month
    		ButtonGroup ButtonPressed
      			OkButton 50, 85, 50, 15
      			CancelButton 105, 85, 50, 15
    		Text 5, 10, 30, 10, "Result:"
    		Text 5, 30, 30, 10, "Method:"
    		Text 5, 50, 30, 10, "Amount:"
    		Text 100, 50, 20, 10, "Date: "
    		Text 110, 70, 30, 10, "Months:"
  		EndDialog

  		DO
    		err_msg = ""
    		Dialog future_savings_dialog
    		cancel_confirmation
    		IF IsNumeric(date_month) = false or len(resolve_time) < 2 THEN err_msg = err_msg & vbNewLine & "* Enter a valid date MM/YY."
    		IF saving_result = "Select One:" THEN err_msg = err_msg & vbNewLine & "Please enter a saving result."
    		IF saving_method = "Select One:" THEN err_msg = err_msg & vbNewLine & "Please enter a saving method."
			IF saving_method = "O - One Time Only" and saving_month <> "" THEN err_msg = err_msg & vbNewLine & "When selecting method O no months need to be entered."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
  		LOOP UNTIL err_msg = ""
	END IF
	''---------------------------------------------------------------------------writing the note on IULB
	EMReadScreen error_msg, 11, 24, 2
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	If error_msg = "ACTION CODE" THEN script_end_procedure(err_msg & vbNewLine & "Please ensure you are selecting the correct code for resolve. PF10 to ensure the match can be resolved using the script.")'checking for error msg'
	IF resolution_status = "BC - Case Closed" 	THEN EMWriteScreen "Case closed. " & other_notes, 8, 6   							'BC
	IF resolution_status = "BE - No Change" THEN EMWriteScreen "No change. " & other_notes, 8, 6 									'BE
	IF resolution_status = "BE - Child" THEN EMWriteScreen "No change, minor child income excluded. " & other_notes, 8, 6 			'BE - child
	IF resolution_status = "BE - OP Entered" THEN EMWriteScreen "OP entered other programs" & other_notes, 8, 6
	IF resolution_status = "BN - Already known, No Savings" THEN EMWriteScreen "Already known - No savings. " & other_notes, 8, 6 	'BN
	IF resolution_status = "BO - Other" THEN EMWriteScreen "HC Claim entered. " & other_notes, 8, 6 								'BO
	IF resolution_status = "BP - Wrong Person" THEN EMWriteScreen "Client name and wage earner name are different. " & other_notes, 8, 6
	IF resolution_status = "CC - Claim Entered" THEN
		EMWriteScreen "Claim entered #" & claim_num & claim_AMT, 8, 6 						 	'CC
		EMWriteScreen  claim_number, 17, 9
	END IF
	IF resolution_status = "CF - Future Savings" THEN
		EMWriteScreen "Cost Savings - income can be budgeted timely for next month", 8, 6 						 	'CF
		EMWriteScreen  Active_Programs, 12, 37
    	EMWriteScreen  saving_result, 12, 42
		EMWriteScreen  saving_method, 12, 49
		EMWriteScreen  saving_amount, 12, 54
		EMWriteScreen  date_month, 12, 65
		EMWriteScreen  date_year, 12, 68
		EMWriteScreen  saving_month, 12, 74
	END IF
	IF resolution_status = "NC - Non Cooperation" THEN EMWriteScreen "Non-coop, requested verification not in ECF, " & other_notes, 8, 6 	'NC
	'msgbox "did the notes input?"
	TRANSMIT 'this will take us back to IEVP main menu'

	'------------------------------------------------------------------back on the IEVP menu, making sure that the match cleared
	EMReadScreen days_pending, 5, 7, 72
	days_pending = trim(days_pending)
	IF IsNumeric(days_pending) = TRUE THEN
		match_cleared = FALSE
		script_end_procedure("This match did not appear to clear. Please check case, and try again.")
	ELSE
		match_cleared = TRUE
	END IF

	'msgbox "Match cleared: " match_cleared
	'IF match_cleared = TRUE THEN
	IEVS_match = replace(IEVS_match, "/", " to ")
	Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days requested for HEADER of casenote'
	PF3 'back to the DAIL'
	'----------------------------------------------------------------the case match CLEARED note
	start_a_blank_CASE_NOTE
		IF IEVS_type = "UBEN" THEN CALL write_variable_in_CASE_NOTE("-----" & UBEN_month & "/" & UBEN_year & " NON-WAGE MATCH (" & type_match & ") " & "(" & first_name & ") CLEARED " & rez_status & "-----")
	  	IF IEVS_type = "BEER" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_year & " NON-WAGE MATCH (" & type_match & ") " & "(" & first_name & ") CLEARED " & rez_status & "-----")
	  	CALL write_bullet_and_variable_in_CASE_NOTE("Period", IEVS_match)
	  	CALL write_bullet_and_variable_in_CASE_NOTE("Active Programs", programs)
		CALL write_bullet_and_variable_in_CASE_NOTE("Source of income", source_income)
		CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- -----")
		IF resolution_status = "BC - Case Closed" 	THEN CALL write_variable_in_CASE_NOTE("* Case closed. ")
		IF resolution_status = "BE - Child" THEN CALL write_variable_in_CASE_NOTE("* Income is excluded for minor child in school.")
		IF resolution_status = "BE - No Change" THEN CALL write_variable_in_CASE_NOTE("* No overpayment or savings related to this match.")
		IF resolution_status = "BE - OP Entered" THEN CALL write_variable_in_CASE_NOTE("* Overpayment or savings were found related to this match.")
		IF resolution_status = "BN - Already known, No Savings" THEN CALL write_variable_in_CASE_NOTE("* Client reported income, correct income in STAT and budgeted.")
		IF resolution_status = "BO - Other" THEN CALL write_variable_in_CASE_NOTE("* HC Claim entered. ")
		IF resolution_status = "BP - Wrong Person" THEN CALL write_variable_in_CASE_NOTE("* Client name and wage earner name are different.  Client's SSN has been verified. No overpayment or savings related to this match.")
		IF resolution_status = "CF - Future Savings" THEN CALL write_variable_in_CASE_NOTE("* Cost Savings - income can be budgeted timely for next month")
		IF resolution_status = "CC - Claim Entered" THEN
			CALL write_variable_in_CASE_NOTE(claim_program & " Overpayment Claim # " & claim_number  & " Amount: $" & claim_AMT &  " From: " & from_month & "/" &  from_year & " through "  & to_month & "/" &  to_year)
			CALL write_bullet_and_variable_in_case_note("Collectible claim", collectible_dropdown)
			CALL write_bullet_and_variable_in_case_note("Reason that claim is collectible or not", collectible_reason)
			CALL write_bullet_and_variable_in_case_note("Verification used for overpayment", EVF_used)
			CALL write_bullet_and_variable_in_case_note("Other responsible member(s)", OT_resp_memb)
			CALL write_bullet_and_variable_in_case_note("Discovery Date", Discovery_date)
			CALL write_bullet_and_variable_in_case_note("Reason for overpayment", reason_OP)
		END IF
		IF resolution_status = "NC - Non Cooperation" THEN
      		CALL write_variable_in_CASE_NOTE("* Client failed to cooperate with wage match.")
      		CALL write_variable_in_case_note("* Entered STAT/DISQ panels for each program.")
      		CALL write_bullet_and_variable_in_case_note("Date Diff notice sent", sent_date)
      		CALL write_variable_in_case_note("* Case approved to close")
      		CALL write_variable_in_case_note("* Client needs to provide: ATR, Income Verification, Difference Notice")
    	END IF
	  	CALL write_bullet_and_variable_in_CASE_NOTE("Responded to Difference Notice", change_response)
	 	CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
    	CALL write_variable_in_CASE_NOTE("----- ----- -----")
	 	CALL write_variable_in_CASE_NOTE ("DEBT ESTABLISHMENT UNIT 612-348-4290 EXT 1-1-1")

		IF TIKL_checkbox = checked THEN
	    'Navigates to TIKL
	    EMSendKey "w"
	    transmit
	    'The following will generate a TIKL formatted date for 10 days from now, and add it to the TIKL
	    CALL create_MAXIS_friendly_date(date, 10, 5, 18)
	    CALL write_variable_in_TIKL("CLOSE FOR IEVS NON-COOP, CREATE DISQ(S) FOR " & first_name)
	    PF3		'Exits and saves TIKL
	    script_end_procedure("Success! Updated NON-WAGE match, and a TIKL created.")
	  END IF
	'END IF
END IF

script_end_procedure ("Match has been acted on. Please take any additional action needed for your case.")
