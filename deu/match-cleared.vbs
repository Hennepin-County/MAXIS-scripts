''GATHERING STATS===========================================================================================
name_of_script = "ACTION-DEU-MATCH-CLEARED.vbs"
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

'Fun with dates! --Creating variables for the rolling 12 calendar months
'current month -1
CM_minus_1_mo =  right("0" &          	 DatePart("m",           DateAdd("m", -1, date)            ), 2)
CM_minus_1_yr =  right(                  DatePart("yyyy",        DateAdd("m", -1, date)            ), 2)
'current month -2'
CM_minus_2_mo =  right("0" &             DatePart("m",           DateAdd("m", -2, date)            ), 2)
CM_minus_2_yr =  right(                  DatePart("yyyy",        DateAdd("m", -2, date)            ), 2)
'current month -3'
CM_minus_3_mo =  right("0" &             DatePart("m",           DateAdd("m", -3, date)            ), 2)
CM_minus_3_yr =  right(                  DatePart("yyyy",        DateAdd("m", -3, date)            ), 2)
'current month -4'
CM_minus_4_mo =  right("0" &             DatePart("m",           DateAdd("m", -4, date)            ), 2)
CM_minus_4_yr =  right(                  DatePart("yyyy",        DateAdd("m", -4, date)            ), 2)
'current month -5'
CM_minus_5_mo =  right("0" &             DatePart("m",           DateAdd("m", -5, date)            ), 2)
CM_minus_5_yr =  right(                  DatePart("yyyy",        DateAdd("m", -5, date)            ), 2)
'current month -6'
CM_minus_6_mo =  right("0" &             DatePart("m",           DateAdd("m", -6, date)            ), 2)
CM_minus_6_yr =  right(                  DatePart("yyyy",        DateAdd("m", -6, date)            ), 2)
'current month -7'
CM_minus_7_mo =  right("0" &             DatePart("m",           DateAdd("m", -7, date)            ), 2)
CM_minus_7_yr =  right(                  DatePart("yyyy",        DateAdd("m", -7, date)            ), 2)
'current month -8'
CM_minus_8_mo =  right("0" &             DatePart("m",           DateAdd("m", -8, date)            ), 2)
CM_minus_8_yr =  right(                  DatePart("yyyy",        DateAdd("m", -8, date)            ), 2)
'current month -9'
CM_minus_9_mo =  right("0" &             DatePart("m",           DateAdd("m", -9, date)            ), 2)
CM_minus_9_yr =  right(                  DatePart("yyyy",        DateAdd("m", -9, date)            ), 2)
'current month -10'
CM_minus_10_mo =  right("0" &            DatePart("m",           DateAdd("m", -10, date)           ), 2)
CM_minus_10_yr =  right(                 DatePart("yyyy",        DateAdd("m", -10, date)           ), 2)
'current month -11'
CM_minus_11_mo =  right("0" &            DatePart("m",           DateAdd("m", -11, date)           ), 2)
CM_minus_11_yr =  right(                 DatePart("yyyy",        DateAdd("m", -11, date)           ), 2)

'Establishing value of variables for the rolling 12 months
current_month = CM_mo & "/" & CM_yr
current_month_minus_one = CM_minus_1_mo & "/" & CM_minus_1_yr
current_month_minus_two = CM_minus_2_mo & "/" & CM_minus_2_yr
current_month_minus_three = CM_minus_3_mo & "/" & CM_minus_3_yr
current_month_minus_four = CM_minus_4_mo & "/" & CM_minus_4_yr
current_month_minus_five = CM_minus_5_mo & "/" & CM_minus_5_yr
current_month_minus_six = CM_minus_6_mo & "/" & CM_minus_6_yr
current_month_minus_seven = CM_minus_7_mo & "/" & CM_minus_7_yr
current_month_minus_eight = CM_minus_8_mo & "/" & CM_minus_8_yr
current_month_minus_nine = CM_minus_9_mo & "/" & CM_minus_9_yr
current_month_minus_ten = CM_minus_10_mo & "/" & CM_minus_10_yr
current_month_minus_eleven = CM_minus_11_mo & "/" & CM_minus_11_yr


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

EMReadscreen dail_check, 4, 2, 48
IF dail_check <> "DAIL" THEN script_end_procedure("You are not in your dail. This script will stop.")
EMSendKey "t"
'checking for an active MAXIS session
Call check_for_MAXIS(FALSE)
EMReadScreen IEVS_type, 4, 6, 6 'read the DAIL msg'
'msgbox IEVS_type
IF IEVS_type = "WAGE" or IEVS_type = "BEER" or IEVS_type = "UBEN" THEN
	match_found = TRUE
ELSE
	script_end_procedure("This is not a IEVS match. Please select a WAGE match DAIL, and run the script again.")
END IF
IF IEVS_type = "BEER" THEN type_match = "B"
IF IEVS_type = "UBEN" THEN type_match = "U"
IF IEVS_type = "WAGE" THEN type_match = "U"

EMReadScreen MAXIS_case_number, 8, 5, 73
MAXIS_case_number= TRIM(MAXIS_case_number)

'----------------------------------------------------------------------------------------------------IEVS
'Navigating deeper into the match interface
CALL write_value_and_transmit("I", 6, 3)   		'navigates to INFC
CALL write_value_and_transmit("IEVP", 20, 71)   'navigates to IEVP
EMReadScreen error_msg, 7, 24, 2
IF error_msg = "NO IEVS" THEN script_end_procedure("An error occurred in IEVP, please process manually.")'checking for error msg'

IF select_quarter = "1" THEN
                IEVS_period = "01-" & CM_minus_1_yr & "/03-" & CM_minus_1_yr
ELSEIF select_quarter = "2" THEN
                IEVS_period = "04-" & CM_minus_1_yr & "/06-" & CM_minus_1_yr
ELSEIF select_quarter = "3" THEN
                IEVS_period = "07-" & CM_minus_1_yr  & "/09-" & CM_minus_1_yr
ELSEIF select_quarter = "4" THEN
                IEVS_period = "10-" & CM_minus_6_yr & "/12-" & CM_minus_6_yr
ELSEIF select_quarter = "YEAR" THEN
				IEVS_period = right(DatePart("yyyy",DateAdd("yyyy", -1, date)), 2)
END IF
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
IF OutOfCounty_error = "MATCH IS NOT" then
	script_end_procedure("Out-of-county case. Cannot update.")
	Else
		IF IEVS_type = "WAGE" then
			EMReadScreen quarter, 1, 8, 14
			EMReadScreen IEVS_year, 4, 8, 22
		ELSEIF IEVS_type = "UBEN" THEN
			EMReadScreen IEVS_month, 2, 5, 68
			EMReadScreen IEVS_year, 4, 8, 71
		ELSEIF IEVS_type = "BEER" THEN
			EMReadScreen IEVS_year, 2, 8, 15
			IEVS_year = "20" & IEVS_year
		END IF
END IF

IF IEVS_type = "BEER" THEN type_match = "B"
IF IEVS_type = "UBEN" THEN type_match = "U"
IF IEVS_type = "WAGE" THEN type_match = "U"
'--------------------------------------------------------------------Client name
EMReadScreen client_name, 35, 5, 24
client_name = trim(client_name)                         'trimming the client name
IF instr(client_name, ",") THEN    						'Most cases have both last name and 1st name. This separates the two names
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
'it is not putting a space in the dialog MsgBox
'TODO change_name_to_FML
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
EMReadScreen source_income, 75, 8, 37
source_income = trim(source_income)
length = len(source_income)		'establishing the length of the variable

IF instr(source_income, " AMOUNT: $") THEN
    position = InStr(source_income, " AMOUNT: $")    		      'sets the position at the deliminator
    source_income = Left(source_income, position)  'establishes employer as being before the deliminator
Elseif instr(source_income, " AMT: $") THEN 					  'establishing the length of the variable
    position = InStr(source_income, " AMT: $")    		      'sets the position at the deliminator
    source_income = Left(source_income, position)  'establishes employer as being before the deliminator
Else
    source_income = source_income	'catch all variable
END IF

'----------------------------------------------------------------------------------------------------Employer info & difference notice info
EMReadScreen notice_sent, 1, 14, 37
EMReadScreen sent_date, 8, 14, 68
sent_date = trim(sent_date)
IF sent_date = "" THEN sent_date = replace(sent_date, "", "N/A")
IF sent_date <> "" THEN sent_date = replace(sent_date, "", "/")

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
  GroupBox 5, 5, 285, 55, "WAGE MATCH"
  Text 10, 20, 110, 10, "Case number: " & MAXIS_case_number
  Text 10, 40, 105, 10, "Active Programs: " & programs
  Text 120, 20, 165, 10, "Client name: " & client_name
  Text 120, 40, 165, 15, "Income source: "  & source_income
  GroupBox 5, 65, 190, 50, "Verification Requested: "
  CheckBox 10, 80, 70, 10, "Difference Notice", Diff_Notice_Checkbox
  CheckBox 110, 80, 90, 10, "Employment Verification", empl_verf_checkbox
  CheckBox 10, 95, 90, 10, "Authorization to Release", ATR_Verf_CheckBox
  CheckBox 110, 95, 80, 10, "Other (please specify)", other_checkbox
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
    cancel_confirmation
    IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL DEU_password_check(False)
	msgbox "ready to go?"
	'--------------------------------------------------------------------sending the notice in IULA
	EMwritescreen "005", 12, 46 'writing the resolve time to read for later
	EMwritescreen "Y", 14, 37 'send Notice
	'msgbox "Difference Notice Sent"
	transmit 'goes into IULA
	'removed the IULB information '
	transmit'exiting IULA, helps prevent errors when going to the case note

	   'Going to the MISC panel
  Call navigate_to_MAXIS_screen ("STAT", "MISC")
  Row = 6
  EmReadScreen panel_number, 1, 02, 78
  If panel_number = "0" then
  	EMWriteScreen "NN", 20,79
  	TRANSMIT
  ELSE
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
    If row = 17 then script_end_procedure("There is not a blank field in the MISC panel. Please delete a line(s), and run script again or update manually.")
  End if

  'writing in the action taken and date to the MISC panel
  EMWriteScreen "Initial Claim Referral", Row, 30
  PF3
	Call navigate_to_MAXIS_screen("DAIL", "WRIT")
	call create_MAXIS_friendly_date(date, 10, 5, 18)
	Call write_variable_in_TIKL("Potential overpayment exists on case. Please review case for receipt of additional requested information.")
  PF3

  start_a_blank_CASE_NOTE
  Call write_variable_in_case_note("----- Claim Referral Tracking -----")
  Call write_bullet_and_variable_in_case_note("Program(s)", program)
  Call write_bullet_and_variable_in_case_note("Action Date", Action_Date)
  Call write_variable_in_case_note("* Entries for these potential claims must be retained until further notice.")
  Call write_variable_in_case_note("---")
  Call write_variable_in_case_note(worker_signature)
	'--------------------------------------------------------------------The case note & case note related code
	pending_verifs = ""
  IF Diff_Notice_Checkbox = CHECKED THEN pending_verifs = pending_verifs & "Difference Notice, "
	IF empl_verf_checkbox = CHECKED THEN pending_verifs = pending_verifs & "EVF, "
	IF ATR_Verf_CheckBox = CHECKED THEN pending_verifs = pending_verifs & "ATR, "
	IF other_checkbox = CHECKED THEN pending_verifs = pending_verifs & "Other, "

    '-------------------------------------------------------------------trims excess spaces of pending_verifs
  pending_verifs = trim(pending_verifs) 	'takes the last comma off of pending_verifs when autofilled into dialog if more than one app date is found and additional app is selected
  IF right(pending_verifs, 1) = "," THEN pending_verifs = left(pending_verifs, len(pending_verifs) - 1)
	IF IEVS_type = "WAGE" THEN
		'Updated IEVS_match to write into case note
		IF quarter = 1 THEN IEVS_quarter = "1ST"
		IF quarter = 2 THEN IEVS_quarter = "2ND"
		IF quarter = 3 THEN IEVS_quarter = "3RD"
		IF quarter = 4 THEN IEVS_quarter = "4TH"
	END IF
	IEVS_match = replace(IEVS_match, "/", " to ")
	Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days

	'---------------------------------------------------------------------DIFF NOTC case note
  start_a_blank_CASE_NOTE
	IF IEVS_type = "WAGE" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH (" & first_name & ") DIFF NOTICE SENT-----")
	IF IEVS_type = "BEER" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_year & " NON-WAGE MATCH (" & type_match & ") " & "(" & first_name & ") DIFF NOTICE SENT-----")
	IF IEVS_type = "UBEN" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_month & " NON-WAGE MATCH (" & type_match & ") " & "(" & first_name & ") DIFF NOTICE SENT-----")
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
	IF sent_date <> "" THEN MsgBox("A difference notice was sent on " & sent_date & "." & vbNewLine & "The script will now navigate to clear the match.")

	BeginDialog cleared_match_dialog, 0, 0, 311, 175, "MATCH CLEARED"
  	Text 10, 20, 110, 10, "Case number: " & MAXIS_case_number
  	Text 120, 20, 165, 10, "Client name: " & client_name
  	Text 10, 40, 105, 10, "Active Programs: " & programs
		Text 120, 40, 175, 15, "Income source: " & source_income
  	DropListBox 75, 65, 110, 15, "Select One: "+chr(9)+"BC - Case Closed"+chr(9)+"BN - Already known, No Savings"+chr(9)+"BE - Child"+chr(9)+"BE - No Change"+chr(9)+"BE - OP Entered"+chr(9)+"BO - Other"+chr(9)+"BP - Wrong Person"+chr(9)+"CC - Claim Entered"+chr(9)+"NC - Non Cooperation", resolution_status
  	DropListBox 125, 85, 60, 15, "Select One: "+chr(9)+"Yes"+chr(9)+"No", change_response
  	EditBox 150, 105, 35, 15, resolve_time
  	EditBox 55, 130, 250, 15, other_notes
  	CheckBox 210, 75, 70, 10, "Difference Notice", Diff_Notice_Checkbox
  	CheckBox 210, 85, 90, 10, "Authorization to Release", ATR_Verf_CheckBox
  	CheckBox 210, 95, 90, 10, "Employment verification", EVF_checkbox
  	CheckBox 210, 105, 80, 10, "Other (please specify)", other_checkbox
  	CheckBox 10, 155, 135, 10, "Check here if 10 day cutoff has passed", TIKL_checkbox
  	ButtonGroup ButtonPressed
  		OkButton 210, 155, 45, 15
			CancelButton 260, 155, 45, 15
		Text 10, 70, 60, 10, "Resolution Status: "
		Text 10, 90, 110, 10, "Responded to Difference Notice: "
		Text 10, 110, 85, 10, "Resolve time (in minutes): "
		Text 10, 135, 40, 10, "Other notes: "
		GroupBox 195, 65, 110, 55, "Verification Used to Clear: "
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
		If (resolution_status = "CC - Claim Entered" AND instr(programs, "HC") or instr(programs, "Medical Assistance")) THEN err_msg = err_msg & vbNewLine & "* System does not allow HC or MA cases to be cleared with the code 'CC - Claim Entered'."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL DEU_password_check(False)

	'----------------------------------------------------------------------------------------------------RESOLVING THE MATCH
	EMWriteScreen resolve_time, 12, 46	    'resolved notes depending on the resolution_status
	IF resolution_status = "BC - Case Closed" THEN rez_status = "BC"
	IF resolution_status = "BE - Child" THEN rez_status = "BE"
	IF resolution_status = "BE - No Change" THEN rez_status = "BE"
	IF resolution_status = "BE - OP Entered" THEN rez_status = "BE"
	IF resolution_status = "BN - Already known, No Savings" THEN rez_status = "BN"
	IF resolution_status = "BO - Other" THEN rez_status = "BO"
	IF resolution_status = "BP - Wrong Person"  THEN rez_status = "BP"
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
	IF resolution_status = "BE - OP Entered" THEN EMWriteScreen "OP entered other programs" & other_notes, 8, 6
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
		script_end_procedure("This match did not appear to clear. Please check case, and try again.")
	ELSE
	  match_cleared = TRUE
	END IF
  IF IEVS_type = "WAGE" THEN
  'Updated IEVS_match to write into case note
  	IF quarter = 1 THEN IEVS_quarter = "1ST"
  	IF quarter = 2 THEN IEVS_quarter = "2ND"
  	IF quarter = 3 THEN IEVS_quarter = "3RD"
  	IF quarter = 4 THEN IEVS_quarter = "4TH"
  END IF
	IEVS_match = replace(IEVS_match, "/", " to ")
	Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days requested for HEADER of casenote'
	PF3 'back to the DAIL'
	   '----------------------------------------------------------------the case match CLEARED note
	start_a_blank_CASE_NOTE
	IF IEVS_type = "WAGE" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH"  & "(" & first_name & ") CLEARED " & rez_status & "-----")
 	IF IEVS_type = "BEER" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_year & " NON-WAGE MATCH (" & type_match & ") " & "(" & first_name & ") CLEARED " & rez_status & "-----")
	IF IEVS_type = "UBEN" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_month & " NON-WAGE MATCH (" & type_match & ") " & "(" & first_name & ") CLEARED " & rez_status & "-----")
	CALL write_bullet_and_variable_in_CASE_NOTE("Period", IEVS_match)
	CALL write_bullet_and_variable_in_CASE_NOTE("Active Programs", programs)
	CALL write_bullet_and_variable_in_CASE_NOTE("Source of income", source_income)
	CALL write_variable_in_CASE_NOTE ("----- ----- -----")
	IF resolution_status = "BC - Case Closed" 	THEN CALL write_variable_in_CASE_NOTE("Case closed. ")
	IF resolution_status = "BE - Child" THEN CALL write_variable_in_CASE_NOTE("INCOME IS EXCLUDED FOR MINOR CHILD IN SCHOOL.")
	IF resolution_status = "BE - OP Entered" THEN CALL write_variable_in_CASE_NOTE("OVERPAYMENTS OR SAVINGS WERE FOUND RELATED TO THIS.")
	IF resolution_status = "BE - No Change" THEN CALL write_variable_in_CASE_NOTE("NO OVERPAYMENTS OR SAVINGS RELATED TO THIS.")
	IF resolution_status = "BN - Already known, No Savings" THEN CALL write_variable_in_CASE_NOTE("CLIENT REPORTED INCOME. CORRECT INCOME IS IN STAT PANELS AND BUDGETED.")
	IF resolution_status = "BO - Other" THEN CALL write_variable_in_CASE_NOTE("HC Claim entered. ")
	IF resolution_status = "BP - Wrong Person" THEN CALL write_variable_in_CASE_NOTE("Client name and wage earner name are different.  Client's SSN has been verified. No overpayment or savings related to this match.")
	IF resolution_status = "CC - Claim Entered" THEN CALL write_variable_in_CASE_NOTE("Claim entered.")
	IF resolution_status = "NC - Non Cooperation" THEN
  	CALL write_variable_in_CASE_NOTE("* CLIENT FAILED TO COOP WITH WAGE MATCH")
    CALL write_variable_in_case_note("* Entered STAT/DISQ panels for each program.")
    CALL write_bullet_and_variable_in_case_note("Date Diff notice sent", sent_date)
    CALL write_variable_in_case_note("* Case approved to close")
    CALL write_variable_in_case_note("* Client needs to provide: ATR, Income Verification, Difference Notice")
  END IF
  CALL write_bullet_and_variable_in_CASE_NOTE("Responded to Difference Notice", change_response)
  CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
  CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- -----")
  CALL write_variable_in_CASE_NOTE ("DEBT ESTABLISHMENT UNIT 612-348-4290 EXT 1-1-1")

	IF TIKL_checkbox = checked THEN    'Navigates to TIKL
	 	EMSendKey "w"
	 	transmit
	 	'The following will generate a TIKL formatted date for 10 days from now, and add it to the TIKL
	 	CALL create_MAXIS_friendly_date(date, 10, 5, 18)
	 	CALL write_variable_in_TIKL("CLOSE FOR IEVS NON-COOP, CREATE DISQ(S) FOR " & first_name)
	 	PF3		'Exits and saves TIKL
	 	script_end_procedure("Success! Updated WAGE match, and a TIKL created.")
	END IF
END IF
script_end_procedure ("Match has been acted on. Please take any additional action needed for your case.")
