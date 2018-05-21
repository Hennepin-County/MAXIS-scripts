'GATHERING STATS===========================================================================================
name_of_script = "OVERPAYMENT CLAIM ENTERED.vbs"
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
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
CALL changelog_update("01/04/2018", "Added handling for HC cases.", "MiKayla Handley, Hennepin County")
CALL changelog_update("01/04/2018", "Added handling for case noting in CCOL.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/11/2017", "Initial version.", "MiKayla Handley, Hennepin County")
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

EMConnect ""
CALL MAXIS_case_number_finder (MAXIS_case_number)
memb_number = "01"
OP_Date = date & ""

BeginDialog EWS_OP_dialog, 0, 0, 396, 220, "Overpayment Claim Entered"
  EditBox 55, 5, 35, 15, MAXIS_case_number
  EditBox 130, 5, 20, 15, memb_number
  EditBox 225, 5, 20, 15, OT_resp_memb
  EditBox 310, 5, 70, 15, Discovery_date
  DropListBox 45, 45, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"HC"+chr(9)+"MF", First_Program
  EditBox 125, 45, 20, 15, First_from_IEVS_month
  EditBox 155, 45, 20, 15, First_from_IEVS_year
  EditBox 195, 45, 20, 15, First_to_IEVS_month
  EditBox 220, 45, 20, 15, First_to_IEVS_year
  EditBox 275, 45, 40, 15, First_OP
  EditBox 340, 45, 40, 15, First_AMT
  DropListBox 45, 65, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"HC"+chr(9)+"MF", Second_Program
  EditBox 125, 65, 20, 15, Second_from_IEVS_month
  EditBox 155, 65, 20, 15, Second_from_IEVS_year
  EditBox 195, 65, 20, 15, Second_to_IEVS_month
  EditBox 220, 65, 20, 15, Second_to_IEVS_year
  EditBox 275, 65, 40, 15, Second_OP
  EditBox 340, 65, 40, 15, Second_AMT
	DropListBox 45, 85, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"HC"+chr(9)+"MF", Third_Program
  EditBox 125, 85, 20, 15, Third_from_IEVS_month
  EditBox 155, 85, 20, 15, Third_from_IEVS_year
  EditBox 195, 85, 20, 15, Third_to_IEVS_month
  EditBox 220, 85, 20, 15, Third_from_IEVS_year
  EditBox 275, 85, 40, 15, Third_OP
  EditBox 340, 85, 40, 15, Third_AMT
  EditBox 80, 180, 40, 15, fed_amount
  EditBox 215, 180, 45, 15, HC_resp_memb
  DropListBox 50, 120, 40, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", collectible_dropdown
  EditBox 165, 120, 120, 15, collectible_reason
  DropListBox 340, 120, 40, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", fraud_referral
  EditBox 60, 140, 80, 15, source_income
  EditBox 235, 140, 145, 15, EVF_used
  EditBox 60, 160, 320, 15, Reason_OP
  CheckBox 265, 185, 120, 10, "Earned Income disregard allowed", EI_checkbox
  ButtonGroup ButtonPressed
    OkButton 285, 200, 45, 15
    CancelButton 335, 200, 45, 15
  Text 5, 10, 50, 10, "Case Number: "
  Text 95, 10, 30, 10, "MEMB #:"
  Text 160, 10, 60, 10, "Other resp. memb:"
  Text 255, 10, 55, 10, "Discovery Date: "
  GroupBox 5, 25, 385, 85, "Overpayment Information"
  Text 10, 50, 30, 10, "Program:"
  Text 100, 50, 20, 10, "From:"
  Text 180, 50, 10, 10, "To:"
  Text 245, 50, 25, 10, "Claim #"
  Text 320, 50, 20, 10, "AMT:"
  Text 10, 70, 30, 10, "Program:"
  Text 100, 70, 20, 10, "From:"
  Text 180, 70, 10, 10, "To:"
  Text 245, 70, 25, 10, "Claim #"
  Text 320, 70, 20, 10, "AMT:"
	Text 10, 90, 30, 10, "Program:"
	Text 100, 90, 20, 10, "From:"
	Text 180, 90, 10, 10, "To:"
	Text 245, 90, 25, 10, "Claim #"
	Text 320, 90, 20, 10, "AMT:"
  Text 5, 185, 70, 10, "Total Fed HC Amount:"
  Text 125, 185, 85, 10, "HC responsible members: "
  Text 5, 125, 40, 10, "Collectible?"
  Text 95, 125, 65, 10, "Collectible Reason:"
  Text 290, 125, 50, 10, "Fraud referral:"
  Text 5, 145, 50, 10, "Income Source: "
  Text 150, 145, 85, 10, "Income verification used:"
  Text 5, 165, 50, 10, "Reason for OP: "
  Text 125, 35, 20, 10, "(MM)"
  Text 155, 35, 15, 10, "(YY)"
  Text 195, 35, 20, 10, "(MM)"
  Text 225, 35, 15, 10, "(YY)"
EndDialog

Do
	err_msg = ""
	dialog EWS_OP_dialog
	IF buttonpressed = 0 THEN stopscript
	IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 THEN err_msg = err_msg & vbnewline & "* Enter a valid case number."
IF First_Program = "Select:"THEN err_msg = err_msg & vbNewLine &  "* Please enter the program for the overpayment."
	IF First_OP = "" THEN err_msg = err_msg & vbnewline & "* You must have an overpayment entry."
	IF First_from_IEVS_month = "Select:" THEN err_msg = err_msg & vbNewLine &  "* Please enter the start month(MM) overpayment occured."
	IF First_from_IEVS_year = "Select:" THEN err_msg = err_msg & vbNewLine &  "* Please enter the start year(YY) overpayment occured."
	IF First_to_IEVS_month = "Select:" THEN err_msg = err_msg & vbNewLine &  "* Please enter the end month(MM) overpayment occured."
	IF First_to_IEVS_year = "Select:" THEN err_msg = err_msg & vbNewLine &  "* Please enter the end year(YY) overpayment occured."
	IF Second_OP <> "" THEN
		IF Second_from_IEVS_month = "Select:" THEN err_msg = err_msg & vbNewLine &  "* Please enter the month(MM) 2nd overpayment occured."
		IF Second_from_IEVS_year = "Select:" THEN err_msg = err_msg & vbNewLine &  "* Please enter the start year(YY) 2nd overpayment occured."
		IF Second_to_IEVS_month = "Select:" THEN err_msg = err_msg & vbNewLine &  "* Please enter the end month(MM) 2nd overpayment occured."
		IF Second_to_IEVS_year = "Select:" THEN err_msg = err_msg & vbNewLine &  "* Please enter the end year(YY) 2nd overpayment occured."
	END IF
	IF Third_OP <> "" THEN
		IF Third_from_IEVS_month = "Select:" THEN err_msg = err_msg & vbNewLine &  "* Please enter the month(MM) 3rd overpayment occured."
		IF Third_from_IEVS_year = "Select:" THEN err_msg = err_msg & vbNewLine &  "* Please enter the start year(YY) 3rd overpayment occured."
		IF Third_to_IEVS_month = "Select:" THEN err_msg = err_msg & vbNewLine &  "* Please enter the end month(MM) 3rd overpayment occured."
		IF Third_to_IEVS_year = "Select:" THEN err_msg = err_msg & vbNewLine &  "* Please enter the end year(YY) 3rd overpayment occured."
	END IF
	IF collectible_dropdown = "Select:"  THEN err_msg = err_msg & vbnewline & "* Please advise if overpayment is collectible."
	IF collectible_dropdown = "YES"  & collectible_reason = "" THEN err_msg = err_msg & vbnewline & "* Please advise why overpayment is collectible."
	IF fraud_referral = "Select:"  THEN err_msg = err_msg & vbnewline & "* Please advise if a fraud referral was made."
	IF source_income = ""  THEN err_msg = err_msg & vbnewline & "* Please advise the source of income."
	IF EVF_used = ""  THEN err_msg = err_msg & vbnewline & "* Please advise the verification used for income."
	IF Reason_OP = ""  THEN err_msg = err_msg & vbnewline & "* Please advise the reason for overpayment."
	IF Discovery_date = ""  THEN err_msg = err_msg & vbnewline & "* Please advise the date the overpayment was discovered (DD/MM/YY)."
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
LOOP UNTIL err_msg = ""
CALL DEU_password_check(False)
'----------------------------------------------------------------------------------------------------Creating the quarter
IF no_match_action = UNCHECKED THEN
    IF select_quarter = "1" THEN
    	IEVS_period = "01-" & CM_yr & "/03-" & CM_yr
    ELSEIF select_quarter = "2" THEN
    	IEVS_period = "04-" & CM_yr & "/06-" & CM_yr
    ELSEIF select_quarter = "3" THEN
        IEVS_period = "07-" & CM_yr & "/09-" & CM_yr
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
CALL write_bullet_and_variable_in_CASE_NOTE("Period", IEVS_period)
CALL write_bullet_and_variable_in_CASE_NOTE("Active Programs", Active_Programs)
IF IEVS_type <> "UNVI" THEN CALL write_bullet_and_variable_in_CASE_NOTE("Source of income", source_income)
Call write_variable_in_CASE_NOTE("----- ----- ----- ----- -----")
Call write_variable_in_CASE_NOTE(other_programs & programs & " Overpayment " & OP_1 & " through " & OP_to_1 & " Claim # " & Claim_1 & " Amt $" & AMT_1)
IF OP_2 <> "" then Call write_variable_in_case_note(other_programs & " Overpayment " & OP_2 & " through  " & OP_to_2 & " Claim # " & Claim_2 & "  Amt $" & AMT_2)
IF OP_3 <> "" then Call write_variable_in_case_note(other_programs & " Overpayment " & OP_3 & " through  " & OP_to_3 & " Claim # " & Claim_3 & "  Amt $" & AMT_3)
IF OP_4 <> "" then Call write_variable_in_case_note(other_programs & " Overpayment " & OP_4 & " through  " & OP_to_4 & " Claim # " & Claim_4 & "  Amt $" & AMT_4)
IF EI_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Earned Income Disregard Allowed")
IF instr(Active_Programs, "HC") then
	Call write_bullet_and_variable_in_CASE_NOTE("HC responsible members", HC_resp_memb)
	Call write_bullet_and_variable_in_CASE_NOTE("Total federal Health Care amount", Fed_HC_AMT)
	Call write_variable_in_CASE_NOTE("---Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
END IF
CALL write_bullet_and_variable_in_case_note("Fraud referral made", fraud_referral)
CALL write_bullet_and_variable_in_case_note("Collectible claim", collectible_dropdown)
CALL write_bullet_and_variable_in_case_note("Reason that claim is collectible or not", collectible_reason)
CALL write_bullet_and_variable_in_case_note("Income verification used", EVF_used)
CALL write_bullet_and_variable_in_case_note("Other responsible member(s)", OT_resp_memb)
CALL write_bullet_and_variable_in_case_note("Reason for overpayment", Reason_OP)
CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- -----")
CALL write_variable_in_CASE_NOTE("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1")
IF instr(Active_Programs, "HC") THEN CALL create_outlook_email("", "mikayla.handley@hennepin.us", "Claims entered for #" &  MAXIS_case_number, "Member #: " & memb_number & vbcr & "Date Overpayment Created: " & OP_Date & vbcr & "Programs: " & program_droplist & vbcr & "See case notes for further details.", "", False)
PF10

Call navigate_to_MAXIS_screen("CCOL", "CLIC")
row = 8			'enters member number
PF9	    'putting EMPS panel into edit mode
Call write_value_and_transmit("x", row, 3)	'opening 'other provider information pop up box
EMReadScreen CCOL_claim, 7, 8, 54
CCOL_claim = trim(CCOL_claim)
msgbox CCOL_claim "& " claim_1
IF CCOL_claim <> Claim_1 THEN
	row = ROW + 1
Else
	Call write_value_and_transmit("x", row, 3)
End if

PF4
msgbox "where are we"

IF IEVS_type = "WAGE" THEN CALL write_variable_in_CCOL_NOTE("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH " & "(" & first_name &  ")" & "OVERPAYMENT-CLAIM ENTERED-----")
IF IEVS_type <> "WAGE" THEN CALL write_variable_in_CCOL_NOTE("-----" & IEVS_year & " NON-WAGE MATCH (" & type_match & ") " & "(" & first_name &  ")" & "OVERPAYMENT-CLAIM ENTERED-----")
CALL write_bullet_and_variable_in_CCOL_NOTE("Period", IEVS_period)
CALL write_bullet_and_variable_in_CCOL_NOTE("Active Programs", Active_Programs)
IF IEVS_type <> "UNVI" THEN CALL write_bullet_and_variable_in_CCOL_NOTE("Source of income", source_income)
Call write_variable_in_CCOL_NOTE("----- ----- ----- ----- -----")
Call write_variable_in_CCOL_NOTE(other_programs & programs & " Overpayment " & OP_1 & " through " & OP_to_1 & " Claim # " & Claim_1 & " Amt $" & AMT_1)
IF OP_2 <> "" then Call write_variable_in_CCOL_NOTE(other_programs & " Overpayment " & OP_2 & " through  " & OP_to_2 & " Claim # " & Claim_2 & "  Amt $" & AMT_2)
IF OP_3 <> "" then Call write_variable_in_CCOL_NOTE(other_programs & " Overpayment " & OP_3 & " through  " & OP_to_3 & " Claim # " & Claim_3 & "  Amt $" & AMT_3)
IF OP_4 <> "" then Call write_variable_in_CCOL_NOTE(other_programs & " Overpayment " & OP_4 & " through  " & OP_to_4 & " Claim # " & Claim_4 & "  Amt $" & AMT_4)
IF EI_checkbox = CHECKED THEN CALL write_variable_in_CCOL_NOTE("* Earned Income Disregard Allowed")
IF instr(Active_Programs, "HC") then
	Call write_bullet_and_variable_in_CCOL_NOTE("HC responsible members", HC_resp_memb)
	Call write_bullet_and_variable_in_CCOL_NOTE("Total federal Health Care amount", Fed_HC_AMT)
	Call write_variable_in_CCOL_NOTE("---Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
END IF
CALL write_bullet_and_variable_in_CCOL_NOTE("Fraud referral made", fraud_referral)
CALL write_bullet_and_variable_in_CCOL_NOTE("Collectible claim", collectible_dropdown)
CALL write_bullet_and_variable_in_CCOL_NOTE("Reason that claim is collectible or not", collectible_reason)
CALL write_bullet_and_variable_in_CCOL_NOTE("Income verification used", EVF_used)
CALL write_bullet_and_variable_in_CCOL_NOTE("Other responsible member(s)", OT_resp_memb)
CALL write_bullet_and_variable_in_CCOL_NOTE("Reason for overpayment", Reason_OP)
CALL write_variable_in_CCOL_NOTE("----- ----- ----- ----- -----")
CALL write_variable_in_CCOL_NOTE("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1")
msgbox "Show me the CCOL"
script_end_procedure("")
