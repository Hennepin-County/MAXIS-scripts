'GATHERING STATS===========================================================================================
name_of_script = "ACTIONS-MATCH-CLEARED-CC.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 180
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
CALL changelog_update("05/14/2018", "Resolved bug preventing 'overpayment reason' from being entered into the case note. Also made this a mandatory field.", "Ilse Ferris, Hennepin County")
CALL changelog_update("05/14/2018", "Fixed bug that prevented script from running. Added End statement, and changed the dialog name.", "Ilse Ferris, Hennepin County")
CALL changelog_update("04/23/2018", "Updated case note to reflect standard dialog and case note.", "MiKayla Handley, Hennepin County")
CALL changelog_update("01/02/2018", "Corrected IEVS match error due to new year.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/14/2017", "Updated script to grab full income source for BEER messages.", "MiKayla Handley, Hennepin County")
CALL changelog_update("06/06/2017", "Initial version.", "MiKayla Handley, Hennepin County")
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

EMConnect ""
CALL MAXIS_case_number_finder (MAXIS_case_number)
MAXIS_case_number= TRIM(MAXIS_case_number)
memb_number = "01"
OP_Date = date & ""

BeginDialog OP_Cleared_dialog, 0, 0, 361, 245, "Overpayment Claim Entered"
	EditBox 55, 5, 35, 15, MAXIS_case_number
	EditBox 150, 5, 45, 15, discovery_date
	EditBox 240, 5, 20, 15, memb_number
  EditBox 330, 5, 20, 15, OT_resp_memb
	DropListBox 55, 25, 40, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", fraud_referral
	DropListBox 150, 25, 50, 15, "Select:"+chr(9)+"BEER"+chr(9)+"UBEN"+chr(9)+"WAGE", IEVS_type
	DropListBox 295, 25, 55, 15, "Select:"+chr(9)+"1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"YEAR"+chr(9)+"LAST YEAR"+chr(9)+"OTHER", select_quarter
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
  EditBox 70, 140, 160, 15, EVF_used
	EditBox 305, 140, 50, 15, income_rcvd_date
	EditBox 70, 160, 285, 15, Reason_OP
	DropListBox 105, 180, 40, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", EI_allowed_dropdown
	DropListBox 195, 180, 40, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", collectible_dropdown
  DropListBox 270, 180, 85, 15, "Select:"+chr(9)+"HH No Info"+chr(9)+"HH Incorrect Or Incompl"+chr(9)+"HH Not Timely Inform"+chr(9)+"Estimate Wrong Inadv"+chr(9)+"No HRF"+chr(9)+"Benefits Rcd Pend App"+chr(9)+"Replacement Bene Used" +chr(9)+"Prog Pol Prevents Chg" +chr(9)+"EBT- Representment" +chr(9)+"IEVS - BEER"+chr(9)+"IEVS - BENDEX" +chr(9)+"IEVS - UNVI"+chr(9)+"IEVS - SDX"+chr(9)+"IEVS - WAGE"+chr(9)+"IEVS - UBEN"+chr(9)+"PARIS Inter-state Match"+chr(9)+"Agency: Delay   Action"+chr(9)+"Agency: Issue Comp Error"+chr(9)+"Agency: Dup Issuance"+chr(9)+"Agency:$50 Pass-Thru"+chr(9)+"GRH Vndr: No Info"+chr(9)+"GRH Vndr: Incomplete"+chr(9)+"GRH Vndr: Not Timely"+chr(9)+"GRH Vndr: Client Left Faci"+chr(9)+"99 Other", collectible_reason_dropdown
	EditBox 95, 200, 60, 15, HC_resp_memb
  EditBox 290, 200, 65, 15, Fed_HC_AMT
  ButtonGroup ButtonPressed
      OkButton 260, 225, 45, 15
      CancelButton 310, 225, 45, 15
  Text 5, 10, 50, 10, "Case Number: "
  Text 95, 10, 55, 10, "Discovery Date: "
  Text 245, 30, 45, 10, "Match Period:"
  Text 5, 30, 50, 10, "Fraud referral:"
  Text 105, 30, 40, 10, "IEVS Type:"
  Text 205, 10, 30, 10, "MEMB #:"
  Text 270, 10, 60, 10, "Other resp. memb:"
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
  Text 150, 185, 40, 10, "Collectible?"
  Text 240, 185, 30, 10, "Reason:"
  Text 5, 145, 60, 10, "Income verif used:"
  Text 05, 205, 90, 10, "HC responsible members:"
  Text 200, 205, 85, 10, "Total Federal HC amount:"
  Text 15, 165, 50, 10, "Reason for OP:"
  Text 240, 145, 60, 10, "Date income rcvd: "
  Text 5, 185, 100, 10, "E. Income disregard allowed?"
	Text 180, 55, 30, 10, "(MM/YY)"
  Text 130, 55, 30, 10, "(MM/YY)"
EndDialog
Do
	err_msg = ""
	dialog OP_Cleared_dialog
	cancel_confirmation
	IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
	IF select_quarter = "Select:" THEN err_msg = err_msg & vbnewline & "* You must select a match period entry."
	IF fraud_referral = "Select:" THEN err_msg = err_msg & vbnewline & "* You must select a fraud referral entry."
	IF IEVS_type = "Select:" THEN err_msg = err_msg & vbnewline & "* You must select a match type entry."
	IF EI_allowed_dropdown = "Select:" THEN err_msg = err_msg & vbnewline & "* Please advise if Earned Income disregard was allowed."
	IF income_rcvd_date = "" THEN err_msg = err_msg & vbnewline & "* Please advise of date income was received."
	IF OP_program = "Select:" THEN err_msg = err_msg & vbnewline & "* You must have an overpayment entry."
  If trim(Reason_OP) = "" THEN err_msg = err_msg & vbnewline & "* You must enter the reason for the overpayment."
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
LOOP UNTIL err_msg = ""
CALL DEU_password_check(False)
'----------------------------------------------------------------------------------------------------Creating the quarter
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
IF edit_error <> "" then script_end_procedure("No IEVS matches and/ or could not access IEVP.")

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
'----------------------------------------------------------------------------------------------------IULA
'Entering the IEVS match & reading the difference notice to ensure this has been sent
EMReadScreen OutOfCounty_error, 12, 24, 2
IF OutOfCounty_error = "MATCH IS NOT" then
	script_end_procedure("Out-of-county case. Cannot update.")
	Else
		IF IEVS_type = "WAGE" then
			EMReadScreen quarter, 1, 8, 14
			EMReadScreen IEVS_year, 4, 8, 22
			If quarter <> select_quarter then script_end_procedure("Match period does not match the selected match period. The script will now end.")
		ELSEIF IEVS_type = "UBEN" THEN
			EMReadScreen IEVS_month, 2, 5, 68
			EMReadScreen IEVS_year, 2, 5, 71
		ELSEIF IEVS_type = "BEER" THEN
			EMReadScreen IEVS_year, 2, 8, 15
			IEVS_year = "20" & IEVS_year
		END IF
END IF

'----------------------------------------------------------------------------------------------------Client name
EMReadScreen client_name, 35, 5, 24
client_name = trim(client_name)                         'trimming the client name
if instr(client_name, ",") then    						'Most cases have both last name and 1st name. This seperates the two names
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

programs = ""
IF instr(Active_Programs, "D") then programs = programs & "DWP, "
IF instr(Active_Programs, "F") then programs = programs & "Food Support, "
IF instr(Active_Programs, "H") then programs = programs & "Health Care, "
IF instr(Active_Programs, "M") then programs = programs & "Medical Assistance, "
IF instr(Active_Programs, "S") then programs = programs & "MFIP, "
'trims excess spaces of programs
programs = trim(programs)
'takes the last comma off of programs when autofilled into dialog
If right(programs, 1) = "," THEN programs = left(programs, len(programs) - 1)

'----------------------------------------------------------------------------------------------------Employer info & diff notice info
EMReadScreen source_income, 74, 8, 37
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

EMReadScreen diff_notice, 1, 14, 37
EMReadScreen diff_date, 10, 14, 68
diff_date = trim(diff_date)
If diff_date <> "" then diff_date = replace(diff_date, " ", "/")

PF3		'exiting IULA, helps prevent errors when going to the case note

'-----------------------------------------------------------------------------------'for the case notes
IF select_quarter = "1" THEN
	IEVS_type = "WAGE"
    IEVS_period = "01-" & CM_yr & "/03-" & CM_yr
ELSEIF select_quarter = "2" THEN
	IEVS_type = "WAGE"
    IEVS_period = "04-" & CM_yr & "/06-" & CM_yr
ELSEIF select_quarter = "3" THEN
    IEVS_type = "WAGE"
	IEVS_period = "07-" & CM_yr & "/09-" & CM_yr
ELSEIF select_quarter = "4" THEN
	IEVS_type = "WAGE"
    IEVS_period = "10-" & CM_minus_6_yr & "/12-" & CM_minus_6_yr
ELSEIF select_quarter = "YEAR" THEN
	IEVS_type = "BEER"
	IEVS_period = right(DatePart("yyyy",DateAdd("yyyy", -1, date)), 2)
END IF

EMWriteScreen "010", 12, 46

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
EMWriteScreen Claim_1, 17, 9
'need to check about adding for mutli claims'
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

IF match_cleared = TRUE THEN
	IF IEVS_type = "WAGE" THEN
        'Updated IEVS_period to write into case note
        IF select_quarter = "1" THEN IEVS_quarter = "1ST"
        IF select_quarter = "2" THEN IEVS_quarter = "2ND"
        IF select_quarter = "3" THEN IEVS_quarter = "3RD"
        IF select_quarter = "4" THEN IEVS_quarter = "4TH"
    END IF
	IEVS_period = replace(IEVS_period, "/", " to ")
	Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days requested for HEADER of casenote'
	PF3 'back to the DAIL'

	'-----------------------------------------------------------------------------------------CASENOTE
	start_a_blank_CASE_NOTE
	IF IEVS_type = "WAGE" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH " & " (" & first_name &  ") " & "OVERPAYMENT-CLAIM ENTERED-----")
	IF IEVS_type = "UBEN" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_month & "/" & IEVS_year & " NON-WAGE MATCH (" & type_match & ") " & "(" & first_name & ") OVERPAYMENT-CLAIM ENTERED-----")
	IF IEVS_type <> "WAGE" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_year & " NON-WAGE MATCH (" & type_match & ") " & " (" & first_name &  ") " & "OVERPAYMENT-CLAIM ENTERED-----")
	CALL write_bullet_and_variable_in_CASE_NOTE("Discovery date", discovery_date)
	CALL write_bullet_and_variable_in_CASE_NOTE("Period", IEVS_period)
	CALL write_bullet_and_variable_in_CASE_NOTE("Active Programs", Active_Programs)
	IF IEVS_type <> "UNVI" THEN CALL write_bullet_and_variable_in_CASE_NOTE("Source of income", source_income)
	Call write_variable_in_CASE_NOTE("----- ----- ----- ----- -----")
	Call write_variable_in_CASE_NOTE(OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number & " Amt $" & Claim_amount)
	IF OP_2 <> "" then Call write_variable_in_CASE_NOTE(OP_program_II & " Overpayment " & OP_from_II & " through " & OP_to_II & " Claim # " & Claim_number_II & " Amt $" & Claim_amount_II)
	IF OP_3 <> "" then Call write_variable_in_CASE_NOTE(OP_program_III & " Overpayment " & OP_from_III & " through " & OP_to_III & " Claim # " & Claim_number_III & " Amt $" & Claim_amount_III)
	CALL write_variable_in_case_note("* Earned Income Disregard Allowed")
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
	IF First_OP_program = "HC" THEN CALL create_outlook_email("HSPH.FIN.Unit.AR.Spaulding@hennepin.us", "mikayla.handley@hennepin.us", "Claim entered for case #" &  MAXIS_case_number, "Member #: " & OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number & " Amt $" & Claim_amount & "See case notes for further details.", "", False)

	script_end_procedure("Overpayment case note entered. Please remember to copy and paste your notes to CCOL/CLIC")
End if
