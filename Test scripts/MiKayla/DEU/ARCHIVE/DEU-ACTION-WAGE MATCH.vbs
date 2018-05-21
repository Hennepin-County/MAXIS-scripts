''GATHERING STATS===========================================================================================
name_of_script = "DEU-ACTION-WAGE MATCH.vbs"
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
CALL changelog_update("11/30/2017", "Added CC - add handling for claim entered", "MiKayla Handley, Hennepin County")
CALL changelog_update("11/27/2017", "Added BP - Wrong Person", "MiKayla Handley, Hennepin County")
CALL changelog_update("11/22/2017", "Updated Noncoop option to the cleared match.", "MiKayla Handley, Hennepin County")
CALL changelog_update("11/21/2017", "Updated to clear match, and added handling for sending the difference notice.", "MiKayla Handley, Hennepin County")
CALL changelog_update("11/14/2017", "Initial version.", "MiKayla Handley, Hennepin County")
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
IF IEVS_type = "WAGE" or IEVS_type = "BEER" THEN
	match_found = TRUE
Else
	script_end_procedure("This is not a IEVS match. Please select a WAGE match DAIL, and run the script again.")
End if

EMReadScreen MAXIS_case_number, 8, 5, 73
MAXIS_case_number= TRIM(MAXIS_case_number)
EMReadScreen DAIL_SSN, 9 , 6, 20


'-----------------------------------------------------------------GOING TO STATFOR MEMB AGE
EMSendKey "s"
transmit
EMReadScreen stat_check, 4, 20, 21
If stat_check <> "STAT" then script_end_procedure("Unable to get to stat due to an error screen. Clear the error screen and return to the DAIL. Then try the script again.")

'GOING TO MEMB, NEED TO CHECK THE HH MEMBER
EMWriteScreen "memb", 20, 71
transmit
DO
	EMReadScreen MEMB_current, 1, 2, 73
	EMReadScreen MEMB_total, 1, 2, 78
	EMReadScreen MEMB_ssn, 11, 7, 42
	IF DAIL_ssn = replace(MEMB_ssn, " ", "") THEN
		EMReadScreen HH_memb, 2, 4, 33
		EMReadScreen memb_age, 2, 8, 76
		IF cint(memb_age) < 19 THEN MsgBox "This client is under 19, ensure appropriate processing of the wage macth."
	END IF
	transmit
LOOP UNTIL (MEMB_current = MEMB_total) or (new_HIRE_SSN = replace(MEMB_ssn, " ", ""))

'Navigates back to DAIL PF3 included
Do
	EMReadScreen DAIL_check, 4, 2, 48
	If DAIL_check = "DAIL" then exit do
	PF3
LOOP UNTIL DAIL_check = "DAIL"


'----------------------------------------------------------------------------------------------------IEVS
'Navigating deeper into the match interface
CALL write_value_and_transmit("I", 6, 3)   		'navigates to INFC 
CALL write_value_and_transmit("IEVP", 20, 71)   'navigates to IEVP
EMReadScreen error_msg, 7, 24, 2
IF error_msg = "NO IEVS" THEN script_end_procedure("An error occurred in IEVP, please process manually.")'checking for error msg'

row = 7
'-------------------------------------------------------------------Ensuring that match has not already been resolved.
Do
	EMReadScreen days_pending, 2, row, 74
	days_pending = trim(days_pending)
	IF IsNumeric(days_pending) = FALSE THEN 
		script_end_procedure("No pending IEVS match found. Please review IEVP.")
	ELSE
		'Entering the IEVS match & reading the difference notice to ensure this has been sent
		EMReadScreen IEVS_period, 11, row, 47
		EMReadScreen start_month, 2, row, 47
		EMReadScreen end_month, 2, row, 53
		IF trim(start_month) = "" or trim(end_month) = "" THEN 
			Found_match = FALSE
		ELSE
			month_difference = abs(end_month) - abs(start_month)
			IF (IEVS_type = "WAGE" and month_difference = 2) THEN 'ensuring if it is a wage the match is a quarter'
				found_match = true
				EXIT DO
			ELSEIF (IEVS_type = "BEER" and month_difference = 11) THEN  'ensuring that if it a beer that the match is a year'
				found_match = True
				EXIT DO
			END IF
		END IF
		row = row + 1
	END IF
LOOP UNTIL row = 17
IF found_match = FALSE THEN script_end_procedure("No pending IEVS match found. Please review IEVP.")
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
	ELSEIF IEVS_type = "BEER" THEN
		EMReadScreen IEVS_year, 2, 8, 15
		IEVS_year = "20" & IEVS_year
	END IF
END IF 

IF IEVS_type = "WAGE" THEN type_match = "U"
IF IEVS_type = "BEER" THEN type_match = "B"

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

'----------------------------------------------------------------------------------------------------Employer info & dIFfernce notice info
EMReadScreen source_income, 27, 8, 37
source_income = trim(source_income)
IF instr(source_income, " AMT: $") THEN 
    length = len(source_income) 						  'establishing the length of the variable
    position = InStr(source_income, " AMT: $")    		      'sets the position at the deliminator  
    source_income = Left(source_income, position)  'establishes employer as being before the deliminator
ELSE 
    source_income = source_income
END IF 

'----------------------------------------------------------------------------------------------------Employer info & difference notice info
EMReadScreen notice_sent, 1, 14, 37
EMReadScreen sent_date, 8, 14, 68
sent_date = trim(sent_date)
IF sent_date <> "" THEN sent_date = replace(sent_date, " ", "/")
IF sent_date = "" THEN sent_date = replace(sent_date, " ", "has not been sent")

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
CALL DEU_password_check(FALSE)

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
	CALL DEU_password_check(FALSE)

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
    
	Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days
	'---------------------------------------------------------------------DIFF NOTC case note
    start_a_blank_CASE_NOTE
    CALL write_variable_in_CASE_NOTE ("-----" & IEVS_year & " WAGE MATCH " & "(" & type_match & ") " & "(" & first_name &  ") DIFF NOTICE SENT-----")
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
    BeginDialog cleared_match_dialog, 0, 0, 311, 175, "WAGE MATCH CLEARED"
      GroupBox 5, 5, 300, 55, "WAGE MATCH"
      Text 10, 20, 110, 10, "Case number:" & MAXIS_case_number
      Text 120, 20, 165, 10, "Client name:" & client_name
      Text 10, 40, 105, 10, "Active Programs:" & programs
      Text 120, 40, 175, 15, "Income source:" & source_income
      DropListBox 75, 65, 110, 15, "Select One:"+chr(9)+"BC - Case Closed"+chr(9)+"BN - Already known, No Savings"+chr(9)+"BE - Child"+chr(9)+"BE - No Change"+chr(9)+"BO - Other"+chr(9)+"BP - Wrong Person"+chr(9)+"CC - Claim Entered"+chr(9)+"NC - Non Cooperation", resolution_status
      DropListBox 125, 85, 60, 15, "Select One:"+chr(9)+"Yes"+chr(9)+"No", change_response
      EditBox 150, 105, 35, 15, resolve_time
      EditBox 55, 130, 250, 15, other_notes
      CheckBox 210, 75, 70, 10, "Difference Notice", Diff_Notice_Checkbox
      CheckBox 210, 85, 90, 10, "Authorization to Release", ATR_Verf_CheckBox
	  CheckBox 210, 105, 80, 10, "Other (please specify)", other_checkbox
      CheckBox 210, 95, 90, 10, "Employment verification", EVF_checkbox
	  Text 10, 70, 60, 10, "Resolution Status: "
      Text 10, 90, 110, 10, "Responded to Difference Notice: "
      Text 10, 110, 85, 10, "Resolve time (in minutes): "
      Text 10, 135, 40, 10, "Other notes: "
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
		IF IsNumeric(resolve_time) = FALSE or len(resolve_time) > 3 THEN err_msg = err_msg & vbNewLine & "* Enter a valid numeric resolved time."
		IF resolve_time = "" THEN err_msg = err_msg & vbNewLine & "Please complete resolve time."
		IF change_response = "Select One:" THEN err_msg = err_msg & vbNewLine & "Did the client respond to Difference Notice?"
		IF resolution_status = "Select One:" THEN err_msg = err_msg & vbNewLine & "Please select a resolution status to continue."
		IF (resolution_status = "BE - No Change" AND other_notes = "") THEN err_msg = err_msg & vbNewLine & "When clearing using BE other notes must be completed."
		If (resolution_status = "CC - Claim Entered" AND instr(programs, "HC") or instr(programs, "Medical Assistance")) THEN err_msg = err_msg & vbNewLine & "* System does not allow HC or MA cases to be cleared with the code 'CC - Claim Entered'."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = ""
		
	CALL DEU_password_check(FALSE)
	
	
	IF resolution_status = "CC - Claim Entered" THEN '-----------------------------------TEST for CC path'
	    CALL MAXIS_case_number_finder (MAXIS_case_number)
	    OP_Date = date & ""
	    BeginDialog CC_Cleared_dialog, 0, 0, 376, 220, "Cleared CC-Claim Entered"
	      EditBox 65, 5, 60, 15, MAXIS_case_number
	      DropListBox 200, 5, 55, 15, "Select One:"+chr(9)+"1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"YEAR", select_quarter
	      EditBox 330, 5, 20, 15, HH_memb
	      DropListBox 45, 40, 55, 15, "Select One:"+chr(9)+"FS"+chr(9)+"MFIP"+chr(9)+"HC"+chr(9)+"GA"+chr(9)+"MSA"+chr(9)+"GRH"+chr(9)+"MFIP HG"+chr(9)+"DWP", program_droplist_1
	      EditBox 130, 35, 35, 15, OP_1
	      EditBox 185, 35, 35, 15, OP_to_1
	      EditBox 255, 35, 35, 15, Claim_1
	      EditBox 320, 35, 45, 15, AMT_1
	      DropListBox 45, 60, 55, 15, "Select One:"+chr(9)+"FS"+chr(9)+"MFIP"+chr(9)+"HC"+chr(9)+"GA"+chr(9)+"MSA"+chr(9)+"GRH"+chr(9)+"MFIP HG"+chr(9)+"DWP", program_droplist_2
	      EditBox 130, 55, 35, 15, OP_2
	      EditBox 185, 55, 35, 15, OP_to_2
	      EditBox 255, 55, 35, 15, Claim_2
	      EditBox 320, 55, 45, 15, Amt_2
	      DropListBox 45, 80, 55, 15, "Select One:"+chr(9)+"FS"+chr(9)+"MFIP"+chr(9)+"HC"+chr(9)+"GA"+chr(9)+"MSA"+chr(9)+"GRH"+chr(9)+"MFIP HG"+chr(9)+"DWP", program_droplist_3
	      EditBox 130, 75, 35, 15, OP_3
	      EditBox 185, 75, 35, 15, OP_to_3
	      EditBox 255, 75, 35, 15, Claim_3
	      EditBox 320, 75, 45, 15, AMT_3
	      DropListBox 45, 100, 55, 15, "Select One:"+chr(9)+"FS"+chr(9)+"MFIP"+chr(9)+"HC"+chr(9)+"GA"+chr(9)+"MSA"+chr(9)+"GRH"+chr(9)+"MFIP HG"+chr(9)+"DWP", program_droplist_4
	      EditBox 130, 95, 35, 15, OP_4
	      EditBox 185, 95, 35, 15, OP_to_4
	      EditBox 255, 95, 35, 15, Claim_4
	      EditBox 320, 95, 45, 15, AMT_4
	      DropListBox 80, 125, 65, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", fraud_referral
	      DropListBox 80, 140, 65, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", EI_allowed
	      EditBox 320, 120, 45, 15, OT_resp_memb
	      EditBox 320, 140, 45, 15, Fed_HC_AMT
	      EditBox 320, 160, 45, 15, HC_resp_memb
	      EditBox 70, 160, 130, 15, source_income_verf
	      EditBox 70, 180, 295, 15, Reason_OP
	      ButtonGroup ButtonPressed
	    	OkButton 270, 200, 45, 15
	    	CancelButton 320, 200, 45, 15
	      Text 10, 10, 50, 10, "Case Number: "
	      Text 150, 10, 45, 10, "Match Period:"
	      Text 295, 10, 35, 10, "MEMB #: "
	      GroupBox 5, 25, 365, 90, "Overpayment Information"
	      Text 10, 40, 30, 10, "Program:"
	      Text 105, 40, 25, 10, "From:"
	      Text 170, 40, 10, 10, "To:"
	      Text 225, 40, 25, 10, "Claim #"
	      Text 295, 40, 20, 10, "AMT:"
	      Text 10, 60, 30, 10, "Program:"
	      Text 105, 60, 20, 10, "From:"
	      Text 170, 60, 10, 10, "To:"
	      Text 225, 60, 25, 10, "Claim #"
	      Text 295, 60, 20, 10, "AMT:"
	      Text 10, 80, 30, 10, "Program:"
	      Text 105, 80, 20, 10, "From:"
	      Text 170, 80, 10, 10, "To:"
	      Text 225, 80, 25, 10, "Claim #"
	      Text 295, 80, 20, 10, "AMT:"
	      Text 10, 100, 30, 10, "Program:"
	      Text 105, 100, 20, 10, "From:"
	      Text 170, 100, 10, 10, "To:"
	      Text 225, 100, 25, 10, "Claim #"
	      Text 295, 100, 20, 10, "AMT:"
	      Text 25, 130, 50, 10, "Fraud referral:"
	      Text 5, 145, 70, 10, "EI disregard allowed:"
	      Text 230, 125, 85, 10, "HC responsible members:"
	      Text 255, 145, 65, 10, "Total FED HC AMT:"
	      Text 225, 165, 95, 10, "Other responsible members:"
	      Text 5, 165, 60, 10, "Income verif used:"
	      Text 15, 185, 50, 10, "Reason for OP: "
	    EndDialog
	    
	    Do
	    	Do
	    		err_msg = ""
	    		dialog CC_Cleared_dialog
	    		IF buttonpressed = 0 then stopscript 
	    		IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = FALSE or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
	    		If (Cleared_status = "CC - Claim Entered" AND instr(programs, "HC") or instr(programs, "Medical Assistance")) then err_msg = err_msg & vbNewLine & "* System does not allow HC or MA cases to be cleared with the code 'CC - Claim Entered'."
	    		IF OP_1 = FALSE THEN err_msg = err_msg & vbnewline & "* You must have an overpayment entry."
	    		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	    	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	    Loop until are_we_passworded_out = FALSE					'loops until user passwords back in
	End if 
	'----------------------------------------------------------------------------------------------------RESOLVING THE MATCH 		
	EMWriteScreen resolve_time, 12, 46
	
	'resolved notes depending on the resolution_status
	IF resolution_status = "BC - Case Closed" THEN rez_status = "BC"  						
	IF resolution_status = "BE - Child" THEN rez_status = "BE"
	IF resolution_status = "BE - No Change" THEN rez_status = "BE"
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
	IF resolution_status = "BN - Already known, No Savings" THEN EMWriteScreen "Already known - No savings. " & other_notes, 8, 6 	'BN
	IF resolution_status = "BO - Other" THEN EMWriteScreen "HC Claim entered. " & other_notes, 8, 6 								'BO
	IF resolution_status = "BP - Wrong Person" THEN EMWriteScreen "Client name and wage earner name are different. " & other_notes, 8, 6 	
	IF resolution_status = "CC - Claim Entered" THEN 
		EMWriteScreen "Claim entered. # " & Claim_1, 8, 6
		IF OP_2 <> "" THEN EMWriteScreen "Claim entered. # " & Claim_2 
		IF OP_3 <> "" THEN EMWriteScreen "Claim entered. # " & Claim_3 
		IF OP_4 <> "" THEN EMWriteScreen "Claim entered. # " & Claim_4	
	END IF
	IF resolution_status = "NC - Non Cooperation" THEN EMWriteScreen "NON-COOP - PAST 10 DAY FOR CLOSURE SET TIKL" & other_notes, 8, 6 	'NC
	msgbox "did the notes input?"
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
    	'Updated IEVS_period to write into case note
    	IF quarter = 1 THEN IEVS_quarter = "1ST"
    	IF quarter = 2 THEN IEVS_quarter = "2ND"
    	IF quarter = 3 THEN IEVS_quarter = "3RD"
    	IF quarter = 4 THEN IEVS_quarter = "4TH"
    END IF
	IEVS_period = replace(IEVS_period, "/", " to ")
	Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days requested for HEADER of casenote'
	PF3 'back to the DAIL
		'-----------------------------------------------------------------------------------------CC CASENOTE
	IF resolution_status = "CC - Claim Entered" THEN 
        start_a_blank_CASE_NOTE
		IF IEVS_type = "WAGE" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH" & "(" & first_name &  ")" & "CLEARED CC-CLAIM ENTERED-----")
		IF IEVS_type = "BEER" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_year & " NON-WAGE MATCH (" & type_match & ") " & "(" & first_name &  ")" &  "CLEARED CC-CLAIM ENTERED-----")
		CALL write_bullet_and_variable_in_CASE_NOTE("Period", IEVS_period)
		CALL write_bullet_and_variable_in_CASE_NOTE("Active Programs", programs)
		CALL write_bullet_and_variable_in_CASE_NOTE("Source of income", source_income)
		Call write_variable_in_CASE_NOTE("----- ----- ----- ----- ----- ----- -----")
		Call write_variable_in_CASE_NOTE(program_droplist_1 & " Overpayment " & OP_1 & " through " & OP_to_1 & "  Claim # " & Claim_1 & "  Amt $" & AMT_1)
		IF OP_2 <> "" then Call write_variable_in_case_note(program_droplist_2 & "Overpayment " & OP_2 & "  through  " & OP_to_2 & "  Claim # " & Claim_2 & "  Amt $" & AMT_2)
		IF OP_3 <> "" then Call write_variable_in_case_note(program_droplist_3 & "Overpayment " & OP_3 & "  through  " & OP_to_3 & "  Claim # " & Claim_3 & "  Amt $" & AMT_3)
		IF OP_4 <> "" then Call write_variable_in_case_note(program_droplist_4 & "Overpayment " & OP_4 & "  through  " & OP_to_4 & "  Claim # " & Claim_4 & "  Amt $" & AMT_4)
		'IF EI_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Earned Income Disregard Allowed")'
		IF instr(program_droplist, "HC") then 
			Call write_bullet_and_variable_in_CASE_NOTE("HC responsible members", HC_resp_memb)
			Call write_bullet_and_variable_in_CASE_NOTE("Total Federal Health Care amount", Fed_HC_AMT)
			Call write_variable_in_CASE_NOTE("---Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
		END IF
		CALL write_bullet_and_variable_in_case_note("Earned Income Disregard Allowed", EI_allowed )
		CALL write_bullet_and_variable_in_case_note("Other responsible member(s)", OT_resp_memb)
		CALL write_bullet_and_variable_in_case_note("Fraud referral made", fraud_referral) 
		CALL write_bullet_and_variable_in_case_note("Reason for overpayment", Reason_OP) 
		CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- ----- ----- -----")
		CALL write_variable_in_CASE_NOTE("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1") 
		IF instr(program_droplist, "HC") THEN CALL create_outlook_email("", "mikayla.handley@hennepin.us", "Claims entered for #" &  MAXIS_case_number, "Member #: " & memb_number & vbcr & "Date Overpayment Created: " & OP_Date & vbcr & "Programs: " & program_droplist & vbcr & "See case notes for further details.", "", FALSE)
    ELSE 
	   '----------------------------------------------------------------the case match CLEARED note (NOT CC)
		start_a_blank_CASE_NOTE
		IF IEVS_type = "WAGE" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH (" & type_match & ") " & "(" & first_name &  ") CLEARED " & rez_status & "-----")
	    IF IEVS_type = "BEER" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_year & " NON-WAGE MATCH (" & type_match & ") " & "(" & first_name &  ") CLEARED " & rez_status & "-----")
	    CALL write_bullet_and_variable_in_CASE_NOTE("Period", IEVS_period)
	   	CALL write_bullet_and_variable_in_CASE_NOTE("Active Programs", programs)
		CALL write_bullet_and_variable_in_CASE_NOTE("Source of income", source_income)
		CALL write_variable_in_CASE_NOTE ("----- ----- -----")
		IF resolution_status = "BC - Case Closed" 	THEN CALL write_variable_in_CASE_NOTE("Case closed. ")
		IF resolution_status = "BE - Child" THEN CALL write_variable_in_CASE_NOTE("INCOME IS EXCLUDED FOR MINOR CHILD IN SCHOOL.")
		IF resolution_status = "BE - No Change" THEN CALL write_variable_in_CASE_NOTE("NO OVERPAYMENTS OR SAVINGS RELATED TO THIS.")
		IF resolution_status = "BN - Already known, No Savings" THEN CALL write_variable_in_CASE_NOTE("CLIENT REPORTED INCOME. CORRECT INCOME IS IN STAT PANELS AND BUDGETED.")
		IF resolution_status = "BO - Other" THEN CALL write_variable_in_CASE_NOTE("HC Claim entered. ")
		IF resolution_status = "BP - Wrong Person" THEN CALL write_variable_in_CASE_NOTE("Client name and wage earner name are different.  Client's SSN has been verified. No overpayment or savings related to this match.")
		IF resolution_status = "NC - Non Cooperation" THEN 
        	CALL write_variable_in_CASE_NOTE("CLIENT FAILED TO COOP WITH NONWAGE MATCH")   
        	CALL write_variable_in_case_note("* Entered STAT/DISQ panels for each program.")
        	CALL write_bullet_and_variable_in_case_note("Date Diff notice sent", sent_date)
        	CALL write_bullet_and_variable_in_case_note("Case approved to close")
        	CALL write_variable_in_case_note("* Client needs to provide: ATR, Income Verification, Difference Notice") 
        END IF
	   	CALL write_bullet_and_variable_in_CASE_NOTE("Responded to Difference Notice", change_response)
	   	CALL write_bullet_and_variable_in_CASE_NOTE("Resolution Status", resolution_status)
		CALL write_bullet_and_variable_in_CASE_NOTE("Other Notes", other_notes)
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
	    	script_end_procedure("Success! Updated WAGE match, and a TIKL created.")
	    END IF
		script_end_procedure ("Success!! Please ensure all required overpayment calculation forms are sent out of ECF." & vbNewLine & "Script does not update panels or navigate to CCOL.")
    END IF
	script_end_procedure ("Match has been cleared. Please take any additional action needed for your case.")   
END IF 
