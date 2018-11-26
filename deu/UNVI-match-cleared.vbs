''GATHERING STATS===========================================================================================
name_of_script = "ACTIONS - DEU-UNVI MATCH CLEARED.vbs"
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
CALL changelog_update("09/10/2018", "Fixed navigation to IEVP bug. Added SSN number read needed to enter IEVP panel.", "Ilse Ferris, Hennepin County")
CALL changelog_update("11/21/2017", "Updated Noncoop option to the cleared match.", "MiKayla Handley, Hennepin County")
CALL changelog_update("11/06/2017", "Updated action to clear the match.", "MiKayla Handley, Hennepin County")
CALL changelog_update("10/24/2017", "Updated action to send difference notice.", "MiKayla Handley, Hennepin County")
CALL changelog_update("10/23/2017", "Initial version.", "MiKayla Handley, Hennepin County")
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
IF IEVS_type <> "UNVI" THEN script_end_procedure("This is not a Non-wage match. Please select a Non-wage match, and run the script again.")
IF IEVS_type = "UNVI" THEN type_match = "U"
EMreadScreen MAXIS_case_number, 8, 5, 73
MAXIS_case_number= TRIM(MAXIS_case_number)
EmReadscreen match_SSN, 9, 6, 20

'----------------------------------------------------------------------------------------------------IEVS

CALL write_value_and_transmit("I", 6, 3)   		'navigates to INFC
CALL write_value_and_transmit("UNVI", 20, 71)   'navigates to UNVI
EMReadScreen error_msg, 7, 24, 2
If error_msg = "NO UNVI RECORD FOUND" THEN script_end_procedure(err_msg & vbNewLine & "An error occurred in UNVI, please process manually - checking for footer month.")'checking for error msg'

'---------------------------------------------------------------------------Reading client name and splitting out the 1st name
EMReadScreen Client_Name, 26, 5, 22
client_name = trim(client_name)                         'trimming the client name
IF instr(client_name, ",") THEN    						'Most cases have both last name and 1st name. This seperates the two names
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
ROW = 8
EMReadScreen IEVS_year_check, 4, row, 6 'Entering the UNVI match & reading the income source
EMReadScreen UNVI_total, 10, row, 11
UNVI_total = trim(UNVI_total)
UNVI_total = replace(UNVI_total, "$", "")

DO
	DO
		unvi_info_confirmation = MsgBox("Press YES to confirm this is the match you wish to act on." & vbNewLine & "For the next match, press NO." & vbNewLine & vbNewLine & _
		"   " & client_name & "  Non-wage match information: " & IEVS_year_check & " for $" & UNVI_total, vbYesNoCancel, "Please confirm this match")
		IF unvi_info_confirmation = vbCancel THEN script_end_procedure ("The script has ended. The match has not been acted on.")
		IF unvi_info_confirmation = vbNo THEN row = row + 1 'ask Ilse about putting in a do to stop the match'
			EMReadScreen IEVS_year_check, 4, row, 6
			IEVS_year_check = trim(IEVS_year_check)
			'msgbox IEVS_year_check
			IF IEVS_year_check = "" THEN script_end_procedure ("The script has ended, no match has not been selected.")
			EMReadScreen UNVI_total, 10, row, 11
			UNVI_total = trim(UNVI_total)
			UNVI_total = replace(UNVI_total, "$", "")
		IF unvi_info_confirmation = vbYes THEN EXIT DO
	LOOP UNTIL unvi_info_confirmation = vbYes
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

EMwritescreen "X", row, 3
TRANSMIT
EMReadScreen source_income, 40, 15, 8
source_income = trim(source_income)
EMReadScreen summary_source, 35, 19, 3
summary_source = trim(summary_source)
IF instr(summary_source, " =") then
    length = len(summary_source) 						  'establishing the length of the variable
    position = InStr(summary_source, " =")    		      'sets the position at the deliminator  
    summary_source = Left(summary_source, position-1)  'establishes employer as being before the deliminator
ELSE
    summary_source = summary_source
END IF

BeginDialog unvi_info_dialog, 0, 0, 216, 135, "NON-WAGE MATCH"
  GroupBox 5, 5, 205, 85, "NON-WAGE MATCH CASE NUMBER "  & MAXIS_case_number
  Text 10, 20, 165, 10, "Client name: "  & client_name
  Text 10, 65, 165, 15, "Income source: "   & summary_source
  Text 10, 35, 165, 15, "Name: "   & source_income
  Text 5, 95, 195, 15, "*PLEASE TAKE NOTE OF INFORMATION - SCRIPT WILL     NOT CASE NOTE INCOME SOURCE OR AMOUNT"
  Text 10, 50, 165, 10, "Total: $"   & UNVI_total
  ButtonGroup ButtonPressed
    OkButton 110, 115, 45, 15
    CancelButton 165, 115, 45, 15
EndDialog

Dialog unvi_info_dialog
IF ButtonPressed = 0 THEN StopScript
CALL DEU_password_check(False)

PF3
PF3
'msgbox "Where am i?"
''----------------------------------------------------------------------going back into IEVP
CALL write_value_and_transmit("I", 6, 3)   		'navigates to INFC
EmWriteScreen match_SSN, 3, 63
CALL write_value_and_transmit("IEVP", 20, 71)   'navigates to IEVP
EMReadScreen error_msg, 7, 24, 2
IF error_msg = "NO IEVS" THEN script_end_procedure("An error occurred in IEVP, please process manually.")'checking for error msg'
ROW = 7
'--------------------------------------------------------------------Ensuring that match has not already been resolved.
Do
	EMReadScreen days_pending, 2, row, 74
	days_pending = trim(days_pending)
	IF IsNumeric(days_pending) = false THEN
		script_end_procedure("No pending IEVS match found. Please review IEVP.")
	ELSE
		'Entering the IEVS match & reading the difference notice to ensure this has been sent
		EMReadScreen IEVS_period, 11, row, 47
		EMReadScreen start_month, 2, row, 47
		EMReadScreen end_month, 2, row, 53
		IF trim(start_month) = "" or trim(end_month) = "" THEN
			Found_match = False
		ELSE
			month_difference = abs(end_month) - abs(start_month)
			IF (IEVS_type = "WAGE" and month_difference = 2) THEN 'ensuring if it is a wage the match is a quarter'
				found_match = true
				EXIT DO
			ELSEIF (IEVS_type = "UNVI" and month_difference = 11) THEN  'ensuring that if it a beer that the match is a year'
				found_match = True
				EXIT DO
			END IF
		END IF
		row = row + 1
	END IF
LOOP UNTIL row = 17
IF found_match = False THEN script_end_procedure("No pending IEVS match found. Please review IEVP.")


DO
	ievp_info_confirmation = MsgBox("Press YES to confirm this is the match you wish to act on." & vbNewLine & "For the next match, press NO." & vbNewLine & vbNewLine & _
	"   " & client_name & "  Non-wage match information: " & IEVS_year_check & " & " & UNVI_total, vbYesNoCancel, "Please ensure the cursor is on the match.")
	IF ievp_info_confirmation = vbCancel THEN script_end_procedure ("The script has ended. The match has not been acted on.")
	IF ievp_info_confirmation = vbNo THEN row = row + 1
	IF ievp_info_confirmation = vbYes THEN EXIT DO
LOOP UNTIL ievp_info_confirmation = vbYes

'---------------------------------------------------------------------IULA
CALL write_value_and_transmit("U", row, 3)
'---------------------------------------------------------------------Reading potential errors for out-of-county cases
EMReadScreen OutOfCounty_error, 12, 24, 2
IF OutOfCounty_error = "MATCH IS NOT" THEN
	script_end_procedure("Out-of-county case. Cannot update.")
ELSE
	IF IEVS_type = "WAGE" THEN
		EMReadScreen IEVS_quarter, 1, 8, 14
		EMReadScreen IEVS_year, 4, 8, 22
	ELSEIF IEVS_type = "UNVI" THEN
		EMReadScreen IEVS_year, 4, 8, 15
		'IEVS_year = "20" & IEVS_year check to make sure it reads the full year
	END IF
END IF

'--------------------------------------------------------------------Client name
EMReadScreen client_name, 26, 5, 24
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

BeginDialog send_notice_dialog, 0, 0, 296, 160, "NON-WAGE MATCH SEND DIFFERENCE NOTICE"
  GroupBox 5, 5, 285, 55, "NON-WAGE MATCH"
  Text 10, 20, 110, 10, "Case number: " & MAXIS_case_number
  Text 10, 40, 105, 10, "Active Programs: " & programs
  Text 120, 20, 165, 10, "Client name: " & client_name
  Text 120, 40, 165, 15, "Income source: "  & source_income
  GroupBox 5, 65, 190, 50, "Verification Requested: "
  CheckBox 10, 80, 70, 10, "Difference Notice", Diff_Notice_Checkbox
  CheckBox 110, 80, 80, 10, "Lottery/Gaming Form", lottery_verf_checkbox
  CheckBox 10, 95, 90, 10, "Authorization to Release", ATR_Verf_CheckBox
  CheckBox 110, 95, 80, 10, "Rental Income Form", rental_checkbox
  Text 5, 125, 40, 10, "Other notes: "
  EditBox 50, 120, 240, 15, other_notes
  ButtonGroup ButtonPressed
	OkButton 195, 140, 45, 15
	CancelButton 245, 140, 45, 15
EndDialog

IF notice_sent = "N" THEN
	DO
    	err_msg = ""
    	Dialog notice_action_dialog
		IF ButtonPressed = 0 THEN StopScript
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
    	IF ButtonPressed = 0 THEN StopScript
    	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL DEU_password_check(False)

	'--------------------------------------------------------------------sending the notice in IULA
	EMwritescreen "005", 12, 46 'writing the resolve time to read for later
	EMwritescreen "Y", 14, 37 'send Notice
	transmit
	'msgbox "notice sent"
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
    '-------------------------------------------------------------------trims excess spaces of pending_verifs
	pending_verifs = ""
    IF Diff_Notice_Checkbox = CHECKED THEN pending_verifs = pending_verifs & "Difference Notice, "
	IF lottery_verf_checkbox = CHECKED THEN pending_verifs = pending_verifs & "Lottery/Gaming Form, "
    IF ATR_Verf_CheckBox = CHECKED THEN pending_verifs = pending_veris & "Authorization to Release, "
    IF rental_checkbox =  CHECKED THEN pending_verifs = pending_verifs & "Rental Income Form, "
    '-------------------------------------------------------------------trims excess spaces of pending_verifs
    pending_verifs = trim(pending_verifs) 	'takes the last comma off of pending_verifs when autofilled into dialog if more more than one app date is found and additional app is selected
	IF right(pending_verifs, 1) = "," THEN pending_verifs = left(pending_verifs, len(pending_verifs) - 1)

	Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days

	'---------------------------------------------------------------------DIFF NOTC case note
    start_a_blank_CASE_NOTE
    CALL write_variable_in_CASE_NOTE ("-----" & IEVS_year & "NON-WAGE MATCH " & "(" & type_match & ") " & "(" & first_name &  ") DIFF NOTICE SENT-----")
    CALL write_bullet_and_variable_in_CASE_NOTE("Client Name", Client_Name)
    CALL write_bullet_and_variable_in_CASE_NOTE("Active Programs", programs)
	CALL write_bullet_and_variable_in_CASE_NOTE("Source of income", source_income)
	CALL write_variable_in_CASE_NOTE("* Type of income:")
    CALL write_variable_in_CASE_NOTE ("----- ----- -----")
    CALL write_bullet_and_variable_in_CASE_NOTE("Verification Requested", pending_verifs)
    CALL write_bullet_and_variable_in_CASE_NOTE("Verification Due", Due_date)
	CALL write_variable_in_CASE_NOTE ("* Client must be provided 10 days to return requested verifications *")
    CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
    CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- ----- ----- -----")
    CALL write_variable_in_CASE_NOTE ("DEBT ESTABLISHMENT UNIT 612-348-4290 EXT 1-1-1")
END IF

IF clear_action_checkbox = CHECKED or notice_sent = "Y" THEN
MsgBox("A difference notice was sent on " & sent_date & "." & vbNewLine & "The script will now navigate to clear the Non-wage match.")
	BeginDialog cleared_match_dialog, 0, 0, 311, 175, "NON-WAGE MATCH CLEARED"
	    GroupBox 5, 5, 300, 55, "NON-WAGE MATCH"
	    Text 10, 20, 110, 10, "Case number: " & MAXIS_case_number
	    Text 120, 20, 165, 10, "Client name: " & client_name
	    Text 10, 40, 105, 10, "Active Programs: " & programs
	    Text 120, 40, 175, 15, "Income source: " & source_income
		DropListBox 75, 65, 110, 15, "Select One:"+chr(9)+"BC - Case Closed"+chr(9)+"BN - Already known, No Savings"+chr(9)+"BE - Child"+chr(9)+"BE - No Change"+chr(9)+"BO - Other"+chr(9)+"CC - Claim Entered"+chr(9)+"NC - Non Cooperation", resolution_status
	    DropListBox 125, 85, 60, 15, "Select One:"+chr(9)+"Yes"+chr(9)+"No", change_response
	    EditBox 150, 105, 35, 15, resolve_time
	    EditBox 60, 130, 245, 15, other_notes
	    CheckBox 210, 75, 70, 10, "Difference Notice", Diff_Notice_Checkbox
	    CheckBox 210, 85, 80, 10, "Rental Income Form", rental_checkbox
	    CheckBox 210, 95, 90, 10, "Authorization to Release", ATR_Verf_CheckBox
	    CheckBox 210, 105, 80, 10, "Lottery/Gaming Form", lottery_verf_checkbox
	    Text 10, 70, 60, 10, "Resolution Status: "
	    Text 10, 90, 110, 10, "Responded to Difference Notice: "
	    Text 10, 110, 85, 10, "Resolve time (in minutes): "
	    Text 15, 135, 40, 10, "Other notes: "
	    GroupBox 195, 65, 110, 55, "Verification Used to Clear: "
		CheckBox 10, 155, 135, 10, "Check here if 10 day cutoff has passed", TIKL_checkbox
	    ButtonGroup ButtonPressed
	      OkButton 210, 155, 45, 15
	      CancelButton 260, 155, 45, 15
	EndDialog

	Do
		err_msg = ""
		Dialog cleared_match_dialog
		IF ButtonPressed = 0 THEN StopScript
		IF IsNumeric(resolve_time) = false or len(resolve_time) > 3 THEN err_msg = err_msg & vbNewLine & "* Enter a valid numeric resolved time."
		IF resolve_time = "" THEN err_msg = err_msg & vbNewLine & "Please complete resolve time."
		IF change_response = "Select One:" THEN err_msg = err_msg & vbNewLine & "Did the client respond to Difference Notice?"
		IF resolution_status = "Select One:" THEN err_msg = err_msg & vbNewLine & "Please select a resolution status to continue."
		IF Active_Programs = "H" and resolution_status = "CC - Claim Entered" THEN err_msg = err_msg & vbNewLine & "CC cannot be used - ACTION CODE FOR ACTH IS INVALID"
		IF Active_Programs = "M" and resolution_status = "CC - Claim Entered" THEN err_msg = err_msg & vbNewLine & "CC cannot be used - ACTION CODE FOR ACTM IS INVALID"
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = ""

	CALL DEU_password_check(False)

	'----------------------------------------------------------------------------------------------------RESOLVING THE MATCH
	EMWriteScreen resolve_time, 12, 46

	'resolved notes depending on the resolution_status
	IF resolution_status = "BC - Case Closed" THEN rez_status = "BC"
	IF resolution_status = "BE - Child" THEN rez_status = "BE"
	IF resolution_status = "BE - No Change" THEN rez_status = "BE"
	IF resolution_status = "BN - Already known, No Savings" THEN rez_status = "BN"
	IF resolution_status = "BO - Other" THEN rez_status = "BO"
	IF resolution_status = "CC - Claim Entered" THEN rez_status = "CC"
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
	'----------------------------------------------------------------------------------------writing the note on IULB
	EMReadScreen error_msg, 11, 24, 2
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	If error_msg = "ACTION CODE" THEN script_end_procedure(err_msg & vbNewLine & "Please ensure you are selecting the correct code for resolve. PF10 to ensure the match can be resolved using the script.")'checking for error msg'

	IF resolution_status = "BC - Case Closed" 	THEN EMWriteScreen "Case closed. " & other_notes, 8, 6   							'BC
	IF resolution_status = "BE - No Change" THEN EMWriteScreen "No change. " & other_notes, 8, 6 									'BE
	IF resolution_status = "BE - Child" THEN EMWriteScreen "No change, minor child income excluded. " & other_notes, 8, 6 			'BE - child
	IF resolution_status = "BN - Already known, No Savings" THEN EMWriteScreen "Already known - No savings. " & other_notes, 8, 6 	'BN
	IF resolution_status = "BO - Other" THEN EMWriteScreen "HC Claim entered. " & other_notes, 8, 6 								'BO
	IF resolution_status = "CC - Claim Entered" THEN EMWriteScreen "Claim entered. " & other_notes, 8, 6 						 	'CC
	IF resolution_status = "NC - Non Cooperation" THEN EMWriteScreen "NON-COOP - PAST 10 DAY FOR CLOSURE SET TIKL" & other_notes, 8, 6 						'NC
	'msgbox "did the notes input?"
	TRANSMIT 'this will take us back to IEVP main menu'

	'------------------------------------------------------------------back on the IEVP menu, making sure that the match cleared
	IF resolution_status = "NC - Non Cooperation" THEN msgbox "CLOSE FOR IEVS NON COOP - CREATE DISQ(S) AND ASSESS POSSIBLE OVERPAYMENTS"
	EMReadScreen days_pending, 5, 7, 72
	days_pending = trim(days_pending)
	If IsNumeric(days_pending) = TRUE then
		match_cleared = FALSE
		script_end_procedure("This match did not appear to clear. Please check case, and try again.")
	Else
		match_cleared = TRUE
	End if

	IF match_cleared = TRUE THEN
	    IEVS_period = replace(IEVS_period, "/", " to ")
		Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days requested for HEADER of casenote'
		PF3 'back to the DAIL'
	   '----------------------------------------------------------------the case match CLEARED note
		start_a_blank_CASE_NOTE
		CALL write_variable_in_CASE_NOTE ("-----" & IEVS_year & " NON-WAGE MATCH (" & type_match & ") " & "(" & first_name &  ") CLEARED " & rez_status & "-----")
		CALL write_bullet_and_variable_in_CASE_NOTE("Client Name", Client_Name)
		CALL write_bullet_and_variable_in_CASE_NOTE("Period", IEVS_period)
		CALL write_bullet_and_variable_in_CASE_NOTE("Active Programs", programs)
		CALL write_bullet_and_variable_in_CASE_NOTE("Source of income", source_income)
		CALL write_variable_in_CASE_NOTE("* Type of income:")
		CALL write_variable_in_CASE_NOTE ("----- ----- -----")
		IF resolution_status = "BN - Already known, No Savings" THEN CALL write_variable_in_CASE_NOTE("CLIENT REPORTED INCOME. CORRECT INCOME IS IN STAT PANELS AND BUDGETED.")
	    IF resolution_status = "BE - Child" THEN CALL write_variable_in_CASE_NOTE("INCOME IS EXCLUDED FOR MINOR CHILD IN SCHOOL.")
		IF resolution_status = "BE - No Change" THEN CALL write_variable_in_CASE_NOTE("NO OVERPAYMENTS OR SAVINGS RELATED TO THIS MATCH.")
		IF resolution_status = "NC - Non Cooperation" THEN
			Call write_variable_in_CASE_NOTE("CLIENT FAILED TO COOP WITH NONWAGE MATCH")
			Call write_variable_in_case_note("* Entered STAT/DISQ panels for each program.")
			Call write_bullet_and_variable_in_case_note("Date Diff notice sent", sent_date)
			Call write_bullet_and_variable_in_case_note("Case approved to close")
			Call write_variable_in_case_note("* Client needs to provide: ATR, Income Verification, Difference Notice")
		END IF
		CALL write_bullet_and_variable_in_CASE_NOTE("Responded to Difference Notice", change_response)
		CALL write_bullet_and_variable_in_CASE_NOTE("Resolution Status", resolution_status)
		CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
		CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- -----")
		CALL write_variable_in_CASE_NOTE ("DEBT ESTABLISHMENT UNIT 612-348-4290 EXT 1-1-1")
	    IF TIKL_checkbox = checked THEN
	    	'Navigates to TIKL
	    	EMSendKey "w"
	    	transmit
	    	'The following will generate a TIKL formatted date for 10 days from now, and add it to the TIKL
	    	CALL create_MAXIS_friendly_date(date, 10, 5, 18)
	    	'EMSetCursor 9, 3		'Setting cursor on 9, 3, because the message goes beyond a single line and EMWriteScreen does not word wrap.
	    	'Sending TIKL text.
	    	CALL write_variable_in_TIKL("CLOSE FOR IEVS NON-COOP, CREATE DISQ(S) FOR " & first_name)
	    	PF3		'Exits and saves TIKL
	    	script_end_procedure("Success! Updated for NON-WAGE match, and a TIKL created.")
	    END IF
	END IF
END IF
script_end_procedure ("NON-COOPERATION FOR NON-WAGE MATCH HAS BEEN UPDATED." & vbnewline & vbnewline & "Please remember to approve case to close and update STAT/DISQ if needed.")
