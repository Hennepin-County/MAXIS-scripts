''GATHERING STATS===========================================================================================
name_of_script = "ACTIONS - DEU-MATCH CLEARED NO DAIL.vbs"
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
CALL changelog_update("01/02/2018", "Corrected IEVS match error due to new year.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/11/2017", "Initial version.", "MiKayla Handley, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'=======================================================================================================END CHANGELOG BLOCK
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
CALL MAXIS_case_number_finder (MAXIS_case_number)
memb_number = "01"
date_recieved = date & ""

'-----------------------------------------------------------------------------------------Initial dialog and do...loop
BeginDialog update_action_dialog, 0, 0, 181, 155, "MATCH CLEARED-NO DAIL"
  EditBox 55, 5, 55, 15, MAXIS_case_number
  EditBox 155, 5, 20, 15, MEMB_Number
  DropListBox 90, 25, 85, 15, "Select One:"+chr(9)+"BC - Case Closed"+chr(9)+"BN - Already known, No Savings"+chr(9)+"BE - Child"+chr(9)+"BE - No Change"+chr(9)+"BE - OP Entered"+chr(9)+"BE - NC Non-collectible"+chr(9)+"BO - Other"+chr(9)+"BP - Wrong Person"+chr(9)+"CC - Claim Entered"+chr(9)+"NC - Non Cooperation", resolution_status
  DropListBox 90, 45, 60, 15, "Select One:"+chr(9)+"1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"YEAR", select_quarter
  DropListBox 90, 65, 60, 15, "Select One:"+chr(9)+"WAGE"+chr(9)+"NON-WAGE", IEVS_type
  DropListBox 90, 85, 85, 15, "Select One:"+chr(9)+"DELETED DISQ"+chr(9)+"PENDING VERF"+chr(9)+"N/A", DISQ_action
  EditBox 50, 105, 125, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 70, 130, 50, 15
    CancelButton 125, 130, 50, 15
  Text 5, 10, 50, 10, "Case Number: "
  Text 120, 10, 30, 10, "MEMB #"
  Text 5, 30, 80, 10, "Update resolution status "
  Text 5, 50, 75, 10, "Match Period (quarter)"
  Text 5, 70, 65, 10, "Wage or Non-Wage"
  Text 5, 90, 75, 10, "DISQ panel addressed"
  Text 5, 110, 45, 10, "Other Notes:"
EndDialog
'enhancment to add this to make the same as match cleared'
'BeginDialog cleared_match_dialog, 0, 0, 311, 175, "MATCH CLEARED"
'	Text 10, 20, 110, 10, "Case number: " & MAXIS_case_number
'	Text 120, 20, 165, 10, "Client name: " & client_name
'	Text 10, 40, 105, 10, "Active Programs: " & programs
'	Text 120, 40, 175, 15, "Income source: " & source_income
'	DropListBox 75, 65, 110, 15, "Select One: "+chr(9)+"BC - Case Closed"+chr(9)+"BN - Already known, No Savings"+chr(9)+"BE - Child"+chr(9)+"BE - No Change"+chr(9)+"BE - OP Entered"+chr(9)+"BO - Other"+chr(9)+"BP - Wrong Person"+chr(9)+"CC - Claim Entered"+chr(9)+"NC - Non Cooperation", resolution_status
'	DropListBox 125, 85, 60, 15, "Select One: "+chr(9)+"Yes"+chr(9)+"No", change_response
'	EditBox 150, 105, 35, 15, resolve_time
'	EditBox 55, 130, 250, 15, other_notes
'	CheckBox 210, 75, 70, 10, "Difference Notice", Diff_Notice_Checkbox
'	CheckBox 210, 85, 90, 10, "Authorization to Release", ATR_Verf_CheckBox
'	CheckBox 210, 95, 90, 10, "Employment verification", EVF_checkbox
'	CheckBox 210, 105, 80, 10, "Other (please specify)", other_checkbox
'	CheckBox 10, 155, 135, 10, "Check here if 10 day cutoff has passed", TIKL_checkbox
'	ButtonGroup ButtonPressed
'		OkButton 210, 155, 45, 15
'		CancelButton 260, 155, 45, 15
'	Text 10, 70, 60, 10, "Resolution Status: "
'	Text 10, 90, 110, 10, "Responded to Difference Notice: "
'	Text 10, 110, 85, 10, "Resolve time (in minutes): "
'	Text 10, 135, 40, 10, "Other notes: "
'	GroupBox 195, 65, 110, 55, "Verification Used to Clear: "
'EndDialog



DO
	err_msg = ""
	dialog update_action_dialog
	cancel_confirmation
	IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
    If (Cleared_status = "CC - Claim Entered" AND instr(programs, "HC") or instr(programs, "Medical Assistance")) then err_msg = err_msg & vbNewLine & "* System does not allow HC or MA cases to be cleared with the code 'CC - Claim Entered'."
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
LOOP UNTIL err_msg = ""									'loops until all errors are resolved

CALL DEU_password_check(False)

'-------------------------------------------------------------------------------------------Defaulting the quarters
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

'msgbox IEVS_period

'----------------------------------------------------------------------------------------------------IEVS

CALL navigate_to_MAXIS_screen("STAT", "MEMB")
EMwritescreen memb_number, 20, 76
transmit

EMReadscreen SSN_number_read, 11, 7, 42
SSN_number_read = replace(SSN_number_read, " ", "")

CALL navigate_to_MAXIS_screen("INFC" , "____")
CALL write_value_and_transmit("IEVP", 20, 71)
CALL write_value_and_transmit(SSN_number_read, 3, 63) '
'----------------------------------------------------------------------------------------------------selecting the correct wage match
Row = 7
DO
''	DO
		EMReadScreen IEVS_match, 11, row, 47
		IF trim(IEVS_match) = "" THEN script_end_procedure("IEVS match for the selected period could not be found. The script will now end.")
		ievp_info_confirmation = MsgBox("Press YES to confirm this is the match you wish to act on." & vbNewLine & "For the next match, press NO." & vbNewLine & vbNewLine & _
		"   " & IEVS_match, vbYesNoCancel, "Please confirm this match")
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
	'LOOP UNTIL IEVS_match = IEVS_period
LOOP UNTIL ievp_info_confirmation = vbYes
'---------------------------------------------------------------------IULA
CALL write_value_and_transmit("U", row, 3)
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

'----------------------------------------------------------------------------------------------------Income info & differnce notice info
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
IF sent_date <> "" THEN sent_date = replace(sent_date, " ", "/")
'----------------------------------------------------------------------------------------------------RESOLVING THE MATCH
EMWriteScreen "005", 12, 46 'write resolve time'

'resolved notes depending on the resolution_status
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
IF resolution_status = "BE - Child" THEN EMWriteScreen "No change, minor child income excluded. " & other_notes, 8, 6 			'BE - child
IF resolution_status = "BE - OP Entered" THEN EMWriteScreen "OP entered other programs" & other_notes, 8, 6
IF resolution_status = "BE - No Change" THEN EMWriteScreen "No change. " & other_notes, 8, 6
IF resolution_status = "BE - NC Non-collectible" THEN EMWriteScreen "Non-Coop remains, but claim is non-collectible ", 8, 6
IF resolution_status = "BN - Already known, No Savings" THEN EMWriteScreen "Already known - No savings. " & other_notes, 8, 6 	'BN
IF resolution_status = "BO - Other" THEN EMWriteScreen "HC Claim entered. " & other_notes, 8, 6 								'BO
IF resolution_status = "BP - Wrong Person" THEN EMWriteScreen "Client name and wage earner name are different. ", 8, 6
IF resolution_status = "CC - Claim Entered" THEN EMWriteScreen "Claim entered. " & other_notes, 8, 6 						 	'CC
IF resolution_status = "NC - Non Cooperation" THEN EMWriteScreen "Non-coop, requested verification not in ECF, " & other_notes, 8, 6 	'NC
'msgbox "did the notes input?"
TRANSMIT 'this will take us back to IEVP main menu'
''------------------------------------------------------------------back on the IEVP menu, making sure that the match cleared
IF IEVS_type = "WAGE" THEN
	'Updated IEVS_period to write into case note
	IF quarter = 1 THEN select_quarter = "1ST"
	IF quarter = 2 THEN select_quarter = "2ND"
	IF quarter = 3 THEN select_quarter = "3RD"
	IF quarter = 4 THEN select_quarter = "4TH"
END IF
IEVS_period = replace(IEVS_period, "/", " to ")
Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days requested for HEADER of casenote'
PF3 'back to the DAIL'
  '----------------------------------------------------------------the case match CLEARED note
start_a_blank_CASE_NOTE
IF IEVS_type = "WAGE" THEN CALL write_variable_in_CASE_NOTE("-----" & select_quarter & " QTR " & IEVS_year & " WAGE MATCH" & "(" & first_name & ") CLEARED " & rez_status & "-----")
IF IEVS_type = "NON-WAGE" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_year & " NON-WAGE MATCH(" & type_match & ") " & "(" & first_name & ") CLEARED " & rez_status & "-----")
CALL write_bullet_and_variable_in_CASE_NOTE("Period", IEVS_period)
CALL write_bullet_and_variable_in_CASE_NOTE("Active Programs", programs)
CALL write_bullet_and_variable_in_CASE_NOTE("Source of income", source_income)
CALL write_variable_in_CASE_NOTE ("----- ----- -----")
IF resolution_status = "BC - Case Closed" 	THEN CALL write_variable_in_CASE_NOTE("Case closed. ")
IF resolution_status = "BE - Child" THEN CALL write_variable_in_CASE_NOTE("INCOME IS EXCLUDED FOR MINOR CHILD IN SCHOOL.")
IF resolution_status = "BE - OP Entered" THEN CALL write_variable_in_CASE_NOTE("OVERPAYMENTS OR SAVINGS WERE FOUND RELATED TO THIS.")
IF resolution_status = "BE - No Change" THEN CALL write_variable_in_CASE_NOTE("NO OVERPAYMENTS OR SAVINGS RELATED TO THIS.")
IF resolution_status = "BE - NC Non-collectible" THEN CALL write_variable_in_CASE_NOTE("NO COLLECTIBLE OVERPAYMENTS RELATED TO THIS MATCH, CLIENT IS STILL NC")
IF resolution_status = "BN - Already known, No Savings" THEN CALL write_variable_in_CASE_NOTE("CLIENT REPORTED INCOME. CORRECT INCOME IS IN STAT PANELS AND BUDGETED.")
IF resolution_status = "BO - Other" THEN CALL write_variable_in_CASE_NOTE("HC Claim entered. ")
IF resolution_status = "BP - Wrong Person" THEN CALL write_variable_in_CASE_NOTE("Client name and wage earner name are different.  Client's SSN has been verified. No overpayment or savings related to this match.")
IF resolution_status = "CC - Claim Entered" THEN CALL write_variable_in_CASE_NOTE("Client name and wage earner name are different.")
IF resolution_status = "NC - Non Cooperation" THEN
	CALL write_variable_in_CASE_NOTE("* CLIENT FAILED TO COOP WITH WAGE MATCH")
	CALL write_variable_in_case_note("* Entered STAT/DISQ panels for each program.")
	CALL write_bullet_and_variable_in_case_note("Date Diff notice sent", sent_date)
	CALL write_variable_in_case_note("* Case approved to close")
	CALL write_variable_in_case_note("* Client needs to provide: ATR, Income Verification, Difference Notice")
END IF
CALL write_bullet_and_variable_in_CASE_NOTE("Responded to Difference Notice", change_response)
CALL write_bullet_and_variable_in_CASE_NOTE("Resolution Status", resolution_status)
CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- -----")
CALL write_variable_in_CASE_NOTE ("DEBT ESTABLISHMENT UNIT 612-348-4290 EXT 1-1-1")


script_end_procedure ("Match has been updated. Please take any additional action needed for your case.")
