''GATHERING STATS===========================================================================================
name_of_script = "ACTION - DEU MATCH CLEARED CC NO DAIL.vbs"
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
CALL changelog_update("07/23/2018", "Updated script to correct version and added case note to email for HC matches and CCOL.", "MiKayla Handley, Hennepin County")
CALL changelog_update("06/15/2018", "Corrected dialog to ensure income information is being stored in notes.", "MiKayla Handley, Hennepin County")
CALL changelog_update("01/02/2018", "Corrected IEVS match error due to new year.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/27/2017", "Updated to handle clearing the match when the date is over 45 days.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/14/2017", "Updated to fix claim entering and case note header.", "MiKayla Handley, Hennepin County")
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

CALL MAXIS_case_number_finder (MAXIS_case_number)
memb_number = "01"
discovery_date = date & ""

back_to_self
'--------------------------------------------------------------------Dialog
BeginDialog OP_Cleared_dialog, 0, 0, 361, 245, "Match Cleared CC Claim Entered-NO DAIL"
  EditBox 55, 5, 45, 15, MAXIS_case_number
  DropListBox 155, 5, 55, 15, "Select:"+chr(9)+"1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"YEAR"+chr(9)+"LAST YEAR"+chr(9)+"OTHER", select_quarter
  EditBox 305, 5, 45, 15, discovery_date
  DropListBox 55, 25, 45, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", fraud_referral
  DropListBox 155, 25, 55, 15, "Select:"+chr(9)+"WAGE"+chr(9)+"BEER", IEVS_type
  EditBox 250, 25, 20, 15, memb_number
  EditBox 330, 25, 20, 15, OT_resp_memb
  DropListBox 50, 65, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"HC"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program
  EditBox 130, 65, 30, 15, OP_from
  EditBox 180, 65, 30, 15, OP_to
  EditBox 245, 65, 35, 15, Claim_number
  EditBox 305, 65, 45, 15, Claim_amount
  DropListBox 50, 85, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"HC"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program_II
  EditBox 130, 85, 30, 15, OP_from_II
  EditBox 180, 85, 30, 15, OP_to_II
  EditBox 245, 85, 35, 15, Claim_number_II
  EditBox 305, 85, 45, 15, Claim_amount_II
  DropListBox 50, 105, 50, 15, "Select:"+chr(9)+"DW"+chr(9)+"FS"+chr(9)+"FG"+chr(9)+"HC"+chr(9)+"GA"+chr(9)+"GR"+chr(9)+"MF"+chr(9)+"MS", OP_program_III
  EditBox 130, 105, 30, 15, OP_from_III
  EditBox 180, 105, 30, 15, OP_to_III
  EditBox 245, 105, 35, 15, Claim_number_III
  EditBox 305, 105, 45, 15, Claim_amount_III
  EditBox 70, 140, 190, 15, collectible_reason
  DropListBox 305, 140, 45, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", collectible_dropdown
  EditBox 70, 160, 160, 15, EVF_used
  EditBox 70, 180, 45, 15, income_rcvd_date
  EditBox 70, 200, 185, 15, Reason_OP
  EditBox 305, 160, 45, 15, HC_resp_memb
  EditBox 305, 180, 45, 15, Fed_HC_AMT
  EditBox 305, 200, 45, 15, hc_claim_number
  CheckBox 70, 225, 120, 10, "Earned Income disregard allowed", EI_checkbox
  ButtonGroup ButtonPressed
    OkButton 255, 225, 45, 15
    CancelButton 305, 225, 45, 15
  Text 5, 10, 50, 10, "Case Number: "
  Text 105, 10, 45, 10, "Match Period:"
  Text 250, 10, 55, 10, "Discovery Date: "
  Text 5, 30, 50, 10, "Fraud referral:"
  Text 110, 30, 40, 10, "IEVS Type:"
  Text 215, 30, 30, 10, "MEMB #:"
  Text 275, 30, 55, 10, "OT resp. memb:"
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
  Text 5, 145, 65, 10, "Collectible Reason:"
  Text 265, 145, 40, 10, "Collectible?"
  Text 5, 165, 60, 10, "Income verif used:"
  Text 240, 165, 65, 10, "HC resp. members:"
  Text 5, 185, 60, 10, "Date income rcvd: "
  Text 240, 185, 65, 10, "Total FED HC AMT:"
  Text 15, 205, 50, 10, "Reason for OP:"
  Text 265, 205, 40, 10, "HC Claim #:"
EndDialog

Do
	err_msg = ""
	dialog OP_Cleared_dialog
	cancel_confirmation
	IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
	IF select_quarter = "Select:" THEN err_msg = err_msg & vbnewline & "* You must select a match period entry."
	IF fraud_referral = "Select:" THEN err_msg = err_msg & vbnewline & "* You must select a fraud referral entry."
	IF trim(Reason_OP) = "" or len(Reason_OP) < 8 THEN err_msg = err_msg & vbnewline & "* You must enter a reason for the overpayment please provide as much detail as possible (min 8)."
	IF OP_program = "Select:"THEN err_msg = err_msg & vbNewLine &  "* Please enter the program for the overpayment."
	IF OP_program_II <> "Select:" THEN
		IF OP_from_II = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the month and year overpayment occurred."
		IF Claim_number_II = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the claim number."
		IF Claim_amount_II = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the amount of claim."
	END IF
	IF OP_program_III <> "Select:" THEN
		IF OP_from_III = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the month and year overpayment occurred."
		IF Claim_number_III = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the claim number."
		IF Claim_amount_III = "" THEN err_msg = err_msg & vbNewLine &  "* Please enter the amount of claim."
	END IF
	IF IEVS_type = "Select:" THEN err_msg = err_msg & vbnewline & "* You must select a match type entry."
	IF EI_allowed_dropdown = "Select:" THEN err_msg = err_msg & vbnewline & "* Please advise if Earned Income disregard was allowed."
  	IF collectible_dropdown = "Select:" THEN err_msg = err_msg & vbnewline & "* Please advise if claim is collectible."
	IF collectible_dropdown = "YES" THEN
	IF collectible_reason_dropdown = "Select:" THEN err_msg = err_msg & vbnewline & "* Please advise why claim is collectible."
	END IF
	IF isdate(income_rcvd_date) = False then err_msg = err_msg & vbNewLine & "* Please enter a valid date for the income recieved."
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
LOOP UNTIL err_msg = ""
CALL DEU_password_check(False)

'----------------------------------------------------------------------------------------------------Getting the ssn for IEVS
CALL navigate_to_MAXIS_screen("STAT", "MEMB")
EMwritescreen memb_number, 20, 76
transmit

EMReadscreen SSN_number_read, 11, 7, 42
SSN_number_read = replace(SSN_number_read, " ", "")

CALL navigate_to_MAXIS_screen("INFC" , "____")
CALL write_value_and_transmit("IEVP", 20, 71)
CALL write_value_and_transmit(SSN_number_read, 3, 63) '
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
IF multiple_match = IEVS_period THEN
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

	'----------------------------------------------------------------------------------------------------RESOLVING THE MATCH
	EMWriteScreen "010", 12, 46
	programs_array = split(programs, ",")
	For each programs in programs_array
		programs = trim(programs)
		IF programs = "DWP" then cleared_header = "ACTD"
		IF programs = "Food Support" then cleared_header = "ACTF"
		IF programs = "Health Care" then cleared_header = "ACTH"
		IF programs = "Medical Assistance" then cleared_header = "ACTM"
		IF programs = "MFIP" then cleared_header = "ACTS"
		row = 11
		col = 57
		EMSearch cleared_header, row, col
		EMWriteScreen "CC", row + 1, col + 1
        'EMwritescreen rez_status, 12, 58
	Next

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
	EMWriteScreen "Claim entered. ", 8, 6
	EMWriteScreen Claim_number, 17, 9
	'need to check about adding for mutli claims'

	'msgbox "did the notes input?"
	TRANSMIT 'this will take us back to IEVP main menu'

	'------------------------------------------------------------------back on the IEVP menu, making sure that the match cleared
	EMReadScreen days_pending, 5, 7, 72
	days_pending = trim(days_pending)
	If IsNumeric(days_pending) = TRUE then
		match_cleared = FALSE
		script_end_procedure("This match did not appear to clear. Please check case, and try again.")
	Else
		match_cleared = TRUE
	End if

	IF match_cleared = TRUE THEN
	    IF IEVS_type = "WAGE" THEN
        	'Updated IEVS_period to write into case note
        	IF select_quarter = 1 THEN IEVS_quarter = "1ST"
        	IF select_quarter = 2 THEN IEVS_quarter = "2ND"
        	IF select_quarter = 3 THEN IEVS_quarter = "3RD"
        	IF select_quarter = 4 THEN IEVS_quarter = "4TH"
        END IF
		IEVS_period = replace(IEVS_period, "/", " to ")
		Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days requested for HEADER of casenote'
		PF3 'back to the DAIL'
        '-----------------------------------------------------------------------------------------CASENOTE
		start_a_blank_CASE_NOTE
	    	IF IEVS_type = "WAGE" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH" & "(" & first_name &  ")" & "CLEARED CC-CLAIM ENTERED-----")
	    	IF IEVS_type = "BEER" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_year & " NON-WAGE MATCH(" & type_match & ") " & "(" & first_name &  ")" &  "CLEARED CC-CLAIM ENTERED-----")
			CALL write_bullet_and_variable_in_CASE_NOTE("Discovery date", discovery_date)
			CALL write_bullet_and_variable_in_CASE_NOTE("Period", IEVS_period)
	    	CALL write_bullet_and_variable_in_CASE_NOTE("Active Programs", programs)
	    	CALL write_bullet_and_variable_in_CASE_NOTE("Source of income", source_income)
			Call write_variable_in_CASE_NOTE("----- ----- ----- ----- -----")
			Call write_variable_in_CASE_NOTE(OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number & " Amt $" & Claim_amount)
			IF OP_program_II <> "Select:" then Call write_variable_in_CASE_NOTE(OP_program_II & " Overpayment " & OP_from_II & " through " & OP_to_II & " Claim # " & Claim_number_II & " Amt $" & Claim_amount_II)
			IF OP_program_III <> "Select:" then Call write_variable_in_CASE_NOTE(OP_program_III & " Overpayment " & OP_from_III & " through " & OP_to_III & " Claim # " & Claim_number_III & " Amt $" & Claim_amount_III)
			IF EI_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Earned Income Disregard Allowed")
			IF programs = "Health Care" or programs = "Medical Assistance" THEN
        	    CALL write_bullet_and_variable_in_CASE_NOTE("HC responsible members", HC_resp_memb)
				Call write_bullet_and_variable_in_CASE_NOTE("HC claim number", hc_claim_number)
        		Call write_bullet_and_variable_in_CASE_NOTE("Total federal Health Care amount", Fed_HC_AMT)
        		Call write_variable_in_CASE_NOTE("---Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
			END IF
        	CALL write_bullet_and_variable_in_case_note("Income verification received", income_rcvd_date)
        	CALL write_bullet_and_variable_in_case_note("Other responsible member(s)", OT_resp_memb)
        	CALL write_bullet_and_variable_in_case_note("Fraud referral made", fraud_referral)
        	CALL write_bullet_and_variable_in_case_note("Collectible claim", collectible_dropdown)
        	CALL write_bullet_and_variable_in_case_note("Reason that claim is collectible or not", collectible_reason)
        	CALL write_bullet_and_variable_in_case_note("Reason for overpayment", Reason_OP)
        	CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- ----- ----- -----")
        	CALL write_variable_in_CASE_NOTE("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1")
        PF3

		IF programs = "Health Care" or programs = "Medical Assistance" THEN
			EMWriteScreen "x", 5, 3
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
		    CALL create_outlook_email("HSPH.FIN.Unit.AR.Spaulding@hennepin.us", "mikayla.handley@hennepin.us","Claims entered for #" &  MAXIS_case_number & " Member # " & memb_number & " Date Overpayment Created: " & OP_Date & "Programs: " & programs, "CASE NOTE" & vbcr & message_array,"", False)
		END IF
	 	'---------------------------------------------------------------writing the CCOL case note'
	'	Call navigate_to_MAXIS_screen("CCOL", "CLSM")
	'	EMWriteScreen Claim_number, 4, 9
	'	Transmit
	'	PF4
	'	IF IEVS_type = "WAGE" THEN CALL write_variable_in_CCOL_NOTE("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH" & "(" & first_name &  ")" & "CLEARED CC-CLAIM ENTERED-----")
	'	IF IEVS_type = "BEER" THEN CALL write_variable_in_CCOL_NOTE("-----" & IEVS_year & " NON-WAGE MATCH(" & type_match & ") " & "(" & first_name &  ")" &  "CLEARED CC-CLAIM ENTERED-----")
	'	CALL write_bullet_and_variable_in_CCOL_NOTE("Discovery date", discovery_date)
	'	CALL write_bullet_and_variable_in_CCOL_NOTE("Period", IEVS_period)
	'	CALL write_bullet_and_variable_in_CCOL_NOTE("Active Programs", programs)
	'	CALL write_bullet_and_variable_in_CCOL_NOTE("Source of income", source_income)
	'	CALL write_variable_in_CCOL_NOTE("----- ----- ----- ----- -----")
	'	CALL write_variable_in_CCOL_NOTE(OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number & " Amt $" & Claim_amount)
	'	IF OP_program_II <> "Select:" then CALL write_variable_in_CCOL_NOTE(OP_program_II & " Overpayment " & OP_from_II & " through " & OP_to_II & " Claim # " & Claim_number_II & " Amt $" & Claim_amount_II)
	'	IF OP_program_III <> "Select:" then CALL write_variable_in_CCOL_NOTE(OP_program_III & " Overpayment " & OP_from_III & " through " & OP_to_III & " Claim # " & Claim_number_III & " Amt $" & Claim_amount_III)
	'	IF EI_checkbox = CHECKED THEN CALL write_variable_in_CCOL_NOTE("* Earned Income Disregard Allowed")
	'	IF programs = "Health Care" or programs = "Medical Assistance" THEN
	'		Call write_bullet_and_variable_in_CCOL_NOTE("HC responsible members", HC_resp_memb)
	'		Call write_bullet_and_variable_in_CCOL_NOTE("HC claim number", hc_claim_number)
	'		Call write_bullet_and_variable_in_CCOL_NOTE("Total federal Health Care amount", Fed_HC_AMT)
	'		CALL write_variable_in_CCOL_NOTE("---Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
	'	END IF
	'	CALL write_bullet_and_variable_in_CCOL_NOTE("Fraud referral made", fraud_referral)
	'	CALL write_bullet_and_variable_in_CCOL_NOTE("Collectible claim", collectible_dropdown)
	'	CALL write_bullet_and_variable_in_CCOL_NOTE("Reason that claim is collectible or not", collectible_reason)
	'	CALL write_bullet_and_variable_in_CCOL_NOTE("Income verification received", income_rcvd_date)
	'	IF OT_resp_memb <> "" THEN CALL write_bullet_and_variable_in_CCOL_NOTE("Other responsible member(s)", OT_resp_memb)
	'	CALL write_bullet_and_variable_in_CCOL_NOTE("MANDATORY-Reason for overpayment", Reason_OP)
	'	CALL write_variable_in_CCOL_NOTE("----- ----- ----- ----- ----- ----- -----")
	'	CALL write_variable_in_CCOL_NOTE("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1")
	'	PF3 'exit the case note'
	'	PF3 'back to dail'
	END IF
script_end_procedure("Overpayment case note entered and copied to CCOL.")
