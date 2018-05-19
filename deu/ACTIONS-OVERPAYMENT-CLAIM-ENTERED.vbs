'GATHERING STATS===========================================================================================
name_of_script = "ACTIONS-OVERPAYMENT-CLAIM-ENTERED.vbs"
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
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
CALL changelog_update("04/02/2018", "Updates to fraud referral for the case note.", "MiKayla Handley, Hennepin County")
CALL changelog_update("02/27/2018", "Added income received date.", "MiKayla Handley, Hennepin County")
CALL changelog_update("01/02/2018", "Corrected IEVS match error due to new year.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/11/2017", "Initial version.", "MiKayla Handley, Hennepin County")
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

'---------------------------------------------------------------------------------------FUNCTION'
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

EMConnect ""
CALL MAXIS_case_number_finder (MAXIS_case_number)
memb_number = "01"
OP_Date = date & ""

BeginDialog OP_Cleared_dialog, 0, 0, 361, 240, "Overpayment Claim Entered"
  EditBox 55, 5, 35, 15, MAXIS_case_number
	EditBox 150, 5, 45, 15, discovery_date
	DropListBox 300, 5, 55, 15, "Select:"+chr(9)+"1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"YEAR"+chr(9)+"LAST YEAR"+chr(9)+"OTHER", select_quarter
  DropListBox 55, 25, 35, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", fraud_referral
	DropListBox 150, 25, 50, 15, "Select:"+chr(9)+"WAGE"+chr(9)+"BEER", IEVS_type
  EditBox 245, 25, 20, 15, memb_number
  EditBox 335, 25, 20, 15, OT_resp_memb
	DropListBox 50, 65, 50, 15, "Select:"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"HC"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"DW", First_OP_program
  EditBox 130, 65, 30, 15, OP_1
  EditBox 180, 65, 30, 15, OP_to_1
  EditBox 245, 65, 35, 15, Claim_1
  EditBox 305, 65, 45, 15, AMT_1
  DropListBox 50, 85, 50, 15, "Select:"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"HC"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"DW", Second_OP_program
  EditBox 130, 85, 30, 15, OP_2
  EditBox 180, 85, 30, 15, OP_to_2
  EditBox 245, 85, 35, 15, Claim_2
  EditBox 305, 85, 45, 15, Amt_2
  DropListBox 50, 105, 50, 15, "Select:"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"HC"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"DW", Third_OP_program
  EditBox 130, 105, 30, 15, OP_3
  EditBox 180, 105, 30, 15, OP_to_3
  EditBox 245, 105, 35, 15, Claim_3
  EditBox 305, 105, 45, 15, AMT_3
  DropListBox 50, 140, 35, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", collectible_dropdown
  EditBox 165, 140, 185, 15, collectible_reason
  EditBox 70, 160, 160, 15, EVF_used
  EditBox 305, 160, 45, 15, HC_resp_memb
	EditBox 70, 180, 45, 15, income_rcvd_date
  EditBox 305, 180, 45, 15, Fed_HC_AMT
	EditBox 70, 200, 280, 15, Reason_OP
	CheckBox 5, 220, 120, 10, "Earned Income disregard allowed", EI_checkbox
  ButtonGroup ButtonPressed
    OkButton 255, 220, 45, 15
    CancelButton 305, 220, 45, 15
  Text 5, 10, 50, 10, "Case Number: "
	Text 95, 10, 55, 10, "Discovery Date: "
  Text 250, 10, 45, 10, "Match Period:"
  Text 5, 30, 50, 10, "Fraud referral:"
  Text 110, 30, 40, 10, "IEVS Type:"
  Text 210, 30, 30, 10, "MEMB #:"
  Text 270, 30, 60, 10, "Other resp. memb:"
  GroupBox 10, 45, 345, 90, "Overpayment Information"
  Text 15, 70, 30, 10, "Program:"
  Text 105, 70, 20, 10, "From:"
  Text 165, 70, 10, 10, "To:"
  Text 215, 70, 25, 10, "Claim #"
  Text 285, 70, 20, 10, "AMT:"
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
  Text 5, 145, 40, 10, "Collectible?"
  Text 95, 145, 65, 10, "Collectible Reason:"
  Text 5, 165, 60, 10, "Income verif used:"
  Text 240, 165, 65, 10, "HC resp. members:"
  Text 240, 185, 65, 10, "Total FED HC AMT:"
	Text 15, 205, 50, 10, "Reason for OP:"
  Text 5, 185, 60, 10, "Date income rcvd: "
EndDialog



Do
	err_msg = ""
	dialog OP_Cleared_dialog
	IF buttonpressed = 0 then stopscript
	IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
	IF select_quarter = "Select:" THEN err_msg = err_msg & vbnewline & "* You must select a match period entry."
	IF fraud_referral = "Select:" THEN err_msg = err_msg & vbnewline & "* You must select a fraud referral entry."
	IF IEVS_type = "Select:" THEN err_msg = err_msg & vbnewline & "* You must select a match type entry."
  IF OP_1 = "" THEN err_msg = err_msg & vbnewline & "* You must have an overpayment entry."
	If trim(Reason_OP) = "" THEN err_msg = err_msg & vbnewline & "* You must enter the reason for the overpayment."
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
LOOP UNTIL err_msg = ""
CALL DEU_password_check(False)
'----------------------------------------------------------------------------------------------------Creating the quarter

'msgbox IEVS_period
IF no_match_action = UNCHECKED THEN
    IF select_quarter = "1" THEN
    	IEVS_period = "01-" & CM_minus_1_yr & "/03-" & CM_minus_1_yr
    ELSEIF select_quarter = "2" THEN
    	IEVS_period = "04-" & CM_minus_1_yr & "/06-" & CM_minus_1_yr
    ELSEIF select_quarter = "3" THEN
        IEVS_period = "07-" & CM_minus_1_yr  & "/09-" & CM_minus_1_yr
    ELSEIF select_quarter = "4" THEN
    	IEVS_period = "10-" & CM_minus_6_yr & "/12-" & CM_minus_6_yr
    ELSEIF select_quarter = "YEAR" THEN
    	IEVS_period = right(DatePart("yyyy",DateAdd("yyyy", -1, date)), 4)
    ELSEIF select_quarter = "LAST YEAR" THEN
    	IEVS_period = right(DatePart("yyyy",DateAdd("yyyy", -2, date)), 4)
    ELSEIF select_quarter = "OTHER" THEN
    	IEVS_period = select_period
    END IF
    '----------------------------------------------------------------------------------------------------IEVS
    CALL navigate_to_MAXIS_screen("STAT", "MEMB")
    EMwritescreen memb_number, 20, 76
    transmit

    EMReadscreen SSN_number_read, 11, 7, 42
    SSN_number_read = replace(SSN_number_read, " ", "")

    CALL navigate_to_MAXIS_screen("INFC" , "____")
    CALL write_value_and_transmit("IEVP", 20, 71)
    CALL write_value_and_transmit(SSN_number_read, 3, 63) '

    EMReadScreen edit_error, 2, 24, 2
    edit_error = trim(edit_error)
    IF edit_error <> "" THEN script_end_procedure("No IEVS matches and/ or could not access IEVP.")
    '---------------------------------------------------------------------------------------------Chosing the match to clear'
	Row = 7
	Do
		EMReadScreen IEVS_match, 11, row, 47
		If trim(IEVS_match) = "" THEN script_end_procedure("IEVS match for the selected period could not be found. The script will now end.")
		If IEVS_match = IEVS_period then
			Exit do
		Else
			row = row + 1
			'msgbox "row: " & row
			If row = 17 then
				PF8
				row = 7
			End if
		End if
	Loop until IEVS_period = select_quarter

    EMReadScreen multiple_match, 11, row + 1, 47
    If multiple_match = IEVS_period then
    	msgbox("More than one match exists for this time period. Determine the match you'd like to act on, and put your cursor in front of that match." & vbcr & "Press OK once match is determined.")
    	EMSendKey "U"
    	transmit
    Else
    	CALL write_value_and_transmit("U", row, 3)   'navigates to IULA
    End if
    '----------------------------------------------------------------------------------------------------IULA
    'Entering the IEVS match & reading the difference notice to ensure this has been sent
    'Reading potential errors for out-of-county cases
    EMReadScreen OutOfCounty_error, 12, 24, 2
    IF OutOfCounty_error = "MATCH IS NOT" then
    	script_end_procedure("Out-of-county case. Cannot update.")
    Else
    	IF IEVS_type = "WAGE" then
    		EMReadScreen quarter, 1, 8, 14
    		EMReadScreen IEVS_year, 4, 8, 22
    		If quarter <> select_quarter then script_end_procedure("Match period does not match the selected match period. The script will now end.")
    	Elseif IEVS_type <> "WAGE" THEN
    		EMReadScreen IEVS_year, 4, 8, 15
    	End if
    End if

    '----------------------------------------------------------------------------------------------------Client name
    EMReadScreen client_name, 35, 5, 24
    'Formatting the client name for the spreadsheet
    client_name = trim(client_name)                         'trimming the client name
    if instr(client_name, ",") then    						'Most cases have both last name and 1st name. This separates the two names
    	length = len(client_name)                           'establishing the length of the variable
    	position = InStr(client_name, ",")                  'sets the position at the deliminator (in this case the comma)
    	last_name = Left(client_name, position-1)           'establishes client last name as being before the deliminator
    	first_name = Right(client_name, length-position)    'establishes client first name as after before the deliminator
    Else                                'In cases where the last name takes up the entire space, then the client name becomes the last name
    	first_name = ""
    	last_name = client_name
    END IF
    if instr(first_name, " ") then   						'If there is a middle initial in the first name, then it removes it
    	length = len(first_name)                        	'trimming the 1st name
    	position = InStr(first_name, " ")               	'establishing the length of the variable
    	first_name = Left(first_name, position-1)       	'trims the middle initial off of the first name
    End if

    '----------------------------------------------------------------------------------------------------ACTIVE PROGRAMS
    EMReadScreen Active_Programs, 13, 6, 68
    Active_Programs = trim(Active_Programs)

    '----------------------------------------------------------------------------------------------------Employer info & diff notice info
    EMReadScreen source_income, 74, 8, 37
    source_income = trim(source_income)
    length = len(source_income)		'establishing the length of the variable

    IF instr(source_income, " AMOUNT: $") THEN
        position = InStr(source_income, " AMOUNT: $")    		      'sets the position at the deliminator
        source_income = Left(source_income, position)  'establishes employer as being before the deliminator
    Elseif instr(source_income, " AMT:") THEN 					  'establishing the length of the variable
        position = InStr(source_income, " AMT: $")    		      'sets the position at the deliminator
        source_income = Left(source_income, position)  'establishes employer as being before the deliminator
    Else
        source_income = source_income	'catch all variable
    END IF

    EMReadScreen diff_notice, 1, 14, 37
    EMReadScreen diff_date, 10, 14, 68
    diff_date = trim(diff_date)
    If diff_date <> "" then diff_date = replace(diff_date, " ", "/")

    IF IEVS_type = "UNVI" THEN source_income = replace(source_income, "")

    PF3		'exiting IULA, helps prevent errors when going to the case note
    '-----------------------------------------------------------------------------------'for the case notes

    programs = ""
    IF instr(Active_Programs, "D") then programs = programs & "DWP, "
    IF instr(Active_Programs, "F") then programs = programs & "Food Support, "
    IF instr(Active_Programs, "H") then programs = programs & "Health Care, "
    IF instr(Active_Programs, "M") then programs = programs & "Medical Assistance, "
    IF instr(Active_Programs, "S") then programs = programs & "MFIP, "

    IF other_programs = CHECKED THEN programs = "Food Support, "
    'trims excess spaces of programs
    programs = trim(programs)
    'takes the last comma off of programs when auto filled into dialog
    IF right(programs, 1) = "," THEN programs = left(programs, len(programs) - 1)
    If IEVS_type = "WAGE" THEN
    	'Updated IEVS_period to write into case note
    	If select_quarter = "1" then IEVS_quarter = "1ST"
    	If select_quarter = "2" then IEVS_quarter = "2ND"
    	If select_quarter = "3" then IEVS_quarter = "3RD"
    	If select_quarter = "4" then IEVS_quarter = "4TH"
    End if
    IF IEVS_type = "UNVI" THEN type_match = "U"
    IF IEVS_type = "BEER" THEN type_match = "B"
    IEVS_period = replace(IEVS_period, "/", " to ")
    Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days requested for HEADER of casenote'
    PF3 'back to the DAIL'
ELSE
    CALL navigate_to_MAXIS_screen("STAT", "MEMB")
    EMwritescreen memb_number, 20, 76
	EMReadScreen first_name, 12, 6, 63
	first_name = trim(first_name)
    transmit
END IF
'-----------------------------------------------------------------------------------------CASENOTE
start_a_blank_CASE_NOTE
IF IEVS_type = "WAGE" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH " & "(" & first_name &  ")" & "OVERPAYMENT-CLAIM ENTERED-----")
IF IEVS_type <> "WAGE" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_year & " NON-WAGE MATCH (" & type_match & ") " & "(" & first_name &  ")" & "OVERPAYMENT-CLAIM ENTERED-----")
CALL write_bullet_and_variable_in_CASE_NOTE("Discovery date", discovery_date)
CALL write_bullet_and_variable_in_CASE_NOTE("Period", IEVS_period)
CALL write_bullet_and_variable_in_CASE_NOTE("Active Programs", Active_Programs)
IF IEVS_type <> "UNVI" THEN CALL write_bullet_and_variable_in_CASE_NOTE("Source of income", source_income)
Call write_variable_in_CASE_NOTE("----- ----- ----- ----- -----")
Call write_variable_in_CASE_NOTE(First_OP_program & " Overpayment " & OP_1 & " through " & OP_to_1 & " Claim # " & Claim_1 & " Amt $" & AMT_1)
IF OP_2 <> "" then Call write_variable_in_case_note(Second_Program & " Overpayment " & OP_2 & " through  " & OP_to_2 & " Claim # " & Claim_2 & "  Amt $" & AMT_2)
IF OP_3 <> "" then Call write_variable_in_case_note(Third_Program & " Overpayment " & OP_3 & " through  " & OP_to_3 & " Claim # " & Claim_3 & "  Amt $" & AMT_3)
IF EI_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Earned Income Disregard Allowed")
IF First_OP_program = "HC" THEN
	Call write_bullet_and_variable_in_CASE_NOTE("HC responsible members", HC_resp_memb)
	Call write_bullet_and_variable_in_CASE_NOTE("Total federal Health Care amount", Fed_HC_AMT)
	Call write_variable_in_CASE_NOTE("---Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
END IF
CALL write_bullet_and_variable_in_case_note("Fraud referral made", fraud_referral)
CALL write_bullet_and_variable_in_case_note("Collectible claim", collectible_dropdown)
CALL write_bullet_and_variable_in_case_note("Reason that claim is collectible or not", collectible_reason)
CALL write_bullet_and_variable_in_case_note("Income verification received", EVF_used)
CALL write_bullet_and_variable_in_case_note("Date income verification was received", income_rcvd_date)
CALL write_bullet_and_variable_in_case_note("Other responsible member(s)", OT_resp_memb)
CALL write_bullet_and_variable_in_case_note("Reason for overpayment", Reason_OP)
CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- -----")
CALL write_variable_in_CASE_NOTE("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1")
PF3
IF First_OP_program = "HC" THEN CALL create_outlook_email("HSPH.FIN.Unit.AR.Spaulding@hennepin.us", "mikayla.handley@hennepin.us", "Claims entered for #" &  MAXIS_case_number, "Member #: " & memb_number & vbcr & "Date Overpayment Created: " & OP_Date & vbcr & "Programs: " & program_droplist & vbcr & "See case notes for further details.", "", False)

script_end_procedure("Overpayment case note entered. Please remember to copy and paste your notes to CCOL/CLIC")
