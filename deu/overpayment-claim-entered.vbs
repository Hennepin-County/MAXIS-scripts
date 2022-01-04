'GATHERING STATS===========================================================================================
name_of_script = "ACTIONS - DEU-OVERPAYMENT CLAIM ENTERED.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 500
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================
'run_locally = TRUE
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
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
CALL changelog_update("10/20/2020", "Removed custom functions from script file. Functions have all been incorporated into the project's Function Library.", "Ilse Ferris, Hennepin County")
CALL changelog_update("01/31/2020", "Removed agency overpayment for cash verbiage.", "MiKayla Handley, Hennepin County")
call changelog_update("08/05/2019", "Updated the term claim referral to use the action taken on MISC.", "MiKayla Handley")
CALL changelog_update("04/15/2019", "Updated script to copy case note to CCOL.", "MiKayla Handley, Hennepin County")
CALL changelog_update("01/30/2019", "Updated script to add areas for multiple claims based on request.", "MiKayla Handley, Hennepin County")
CALL changelog_update("04/02/2018", "Updates to fraud referral for the case note.", "MiKayla Handley, Hennepin County")
CALL changelog_update("07/23/2018", "Updated script to correct version and added case note to email for HC matches.", "MiKayla Handley, Hennepin County")
CALL changelog_update("04/02/2018", "Updates to fraud referral for the case note.", "MiKayla Handley, Hennepin County")
CALL changelog_update("02/27/2018", "Added income received date.", "MiKayla Handley, Hennepin County")
CALL changelog_update("01/02/2018", "Corrected IEVS match error due to new year.", "MiKayla Handley, Hennepin County")
CALL changelog_update("12/11/2017", "Initial version.", "MiKayla Handley, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'---------------------------------------------------------------------------------------The script
'Grabs the case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)
memb_number = "01" 'defaults to 01'
discovery_date = "" & date
transfer_to_worker = "X127720" ' add the transfer to the end of script '


DO
'Drop down list for available household members to put into the drop down list.
	Call HH_member_custom_dialog(HH_Member_Array)
	Call convert_array_to_droplist_items(HH_Member_Array, hh_member_dropdown)
	IF uBound(HH_member_array) = -1 THEN MsgBox ("You must select at least one person.")
LOOP UNTIL uBound(HH_member_array) <> -1

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 311, 180, "Claims/Overpayment"
  EditBox 60, 20, 45, 15, MAXIS_case_number
  EditBox 170, 20, 45, 15, discovery_date
  EditBox 280, 20, 25, 15, OT_resp_memb
  EditBox 60, 40, 45, 15, Claim_number
  EditBox 110, 40, 45, 15, Claim_number_II
  EditBox 160, 40, 45, 15, claim_number_III
  EditBox 210, 40, 45, 15, OP_from_IV
  EditBox 260, 40, 45, 15, OP_from_V
  EditBox 95, 60, 45, 15, income_rcvd_date
  EditBox 95, 80, 210, 15, EVF_used
  EditBox 65, 100, 240, 15, Reason_OP
  EditBox 65, 120, 130, 15, verif_requested
  EditBox 65, 140, 135, 15, other_notes
  EditBox 65, 160, 100, 15, worker_signature
  CheckBox 190, 65, 120, 10, "Earned income disregard allowed", EI_checkbox
  DropListBox 260, 120, 45, 15, "Select:"+chr(9)+"YES"+chr(9)+"NO", fraud_referral
  ButtonGroup ButtonPressed
    PushButton 220, 140, 85, 15, "Claims Procedures", claims_procedures_btn
    OkButton 210, 160, 45, 15
    CancelButton 260, 160, 45, 15
  Text 5, 5, 295, 10, "This script will only enter/update STAT/MISC panel for a SNAP or MFIP federal food claim. "
  Text 5, 25, 50, 10, "Case Number: "
  Text 115, 25, 55, 10, "Discovery Date: "
  Text 220, 25, 60, 10, "OT resp. Memb #:"
  Text 5, 45, 55, 10, "Claim Number: "
  Text 5, 65, 85, 10, "Date Income Received: "
  Text 5, 85, 85, 10, "Income Verification Used:"
  Text 5, 105, 40, 10, "OP Reason:"
  Text 5, 125, 55, 10, "Verif Requested:"
  Text 5, 145, 45, 10, "Other Notes:"
  Text 5, 165, 40, 10, "Worker Sig:"
  Text 210, 125, 50, 10, "Fraud referral:"
EndDialog

Do
	err_msg = ""
	dialog Dialog1
	cancel_without_confirmation
	IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
    IF select_quarter = "Select:" THEN err_msg = err_msg & vbnewline & "* You must select a match period entry."
	IF fraud_referral = "Select:" THEN err_msg = err_msg & vbnewline & "* You must select a fraud referral entry."
	IF trim(Reason_OP) = "" or len(Reason_OP) < 5 THEN err_msg = err_msg & vbnewline & "* You must enter a reason for the overpayment please provide as much detail as possible (min 5)."
    IF EVF_used = "" then err_msg = err_msg & vbNewLine & "* Please enter verification used for the income received. If no verification was received enter N/A."
	IF income_rcvd_date <> "" THEN
	 	isdate(income_rcvd_date) = False then err_msg = err_msg & vbNewLine & "* Please enter a valid date for the income received."
	END IF
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
LOOP UNTIL err_msg = ""
CALL check_for_password_without_transmit(are_we_passworded_out)

back_to_self
'----------------------------------------------------------------------------------------------------STAT
CALL navigate_to_MAXIS_screen("STAT", "MEMB")
EMwritescreen memb_number, 20, 76
EMReadScreen first_name, 12, 6, 63
first_name = replace(first_name, "_", "")
first_name = trim(first_name)
transmit
EMReadscreen SSN_number_read, 11, 7, 42
SSN_number_read = replace(SSN_number_read, " ", "")
CALL navigate_to_MAXIS_screen("INFC" , "____")
CALL write_value_and_transmit("IEVP", 20, 71)
CALL write_value_and_transmit(SSN_number_read, 3, 63) '
EMReadScreen edit_error, 2, 24, 2
edit_error = trim(edit_error)
IF edit_error <> "" THEN script_end_procedure_with_error_report("No IEVS matches and/ or could not access IEVP.")

'---------------------------------------------------------------------------------------------Chosing the match to clear'
Row = 7
DO
	EMReadScreen IEVS_match, 11, row, 47
	IF trim(IEVS_match) = "" THEN script_end_procedure_with_error_report("IEVS match for the selected period could not be found. The script will now end.")
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
	IF ievp_info_confirmation = vbCancel THEN script_end_procedure_with_error_report("The script has ended. The match has not been acted on.")
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

'----------------------------------------------------------------------------------------------------IULA
'Entering the IEVS match & reading the difference notice to ensure this has been sent
'Reading potential errors for out-of-county cases
EMReadScreen OutOfCounty_error, 12, 24, 2
IF OutOfCounty_error = "MATCH IS NOT" then
	script_end_procedure_with_error_report("Out-of-county case. Cannot update.")
Else
	IF IEVS_type = "WAGE" then
		EMReadScreen quarter, 1, 8, 14
		EMReadScreen IEVS_year, 4, 8, 22
		If quarter <> select_quarter then script_end_procedure_with_error_report("Match period does not match the selected match period. The script will now end.")
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
'----------------------------------------------------------------------------------------------------Employer info & diff notice info
EMReadScreen income_source, 74, 8, 37
income_source = trim(income_source)
length = len(income_source)		'establishing the length of the variable
IF instr(income_source, " AMOUNT: $") THEN
    position = InStr(income_source, " AMOUNT: $")    		      'sets the position at the deliminator
    income_source = Left(income_source, position)  'establishes employer as being before the deliminator
Elseif instr(income_source, " AMT:") THEN 					  'establishing the length of the variable
    position = InStr(income_source, " AMT: $")    		      'sets the position at the deliminator
    income_source = Left(income_source, position)  'establishes employer as being before the deliminator
Else
    income_source = income_source	'catch all variable
END IF
EMReadScreen diff_notice, 1, 14, 37
EMReadScreen diff_date, 10, 14, 68
diff_date = trim(diff_date)
If diff_date <> "" then diff_date = replace(diff_date, " ", "/")
IF IEVS_type = "UNVI" THEN income_source = replace(income_source, "")
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


'---------------------------------------------------------------writing the CCOL case note'
msgbox "Navigating to CCOL to add case note, please contact the BlueZone Scripts team with any concerns."
Call navigate_to_MAXIS_screen("CCOL", "CLSM")
EMWriteScreen Claim_number, 4, 9
TRANSMIT
'NO CLAIMS WERE FOUND FOR THIS CASE, PROGRAM, AND STATUS
EMReadScreen error_check, 75, 24, 2	'making sure we can actually update this case.
error_check = trim(error_check)
If error_check <> "" then script_end_procedure_with_error_report(error_check & ". Unable to update this case. Please review case, and run the script again if applicable.")
PF4
EMReadScreen existing_case_note, 1, 5, 6
IF existing_case_note = "" THEN
	PF4
ELSE
	PF9
END IF
'-----------------------------------------------------------------------------------------CCOL CASENOTE
IF IEVS_type = "WAGE" THEN CALL write_variable_in_CCOL_note("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH"  & "(" & first_name & ") CLEARED CC-CLAIM ENTERED-----")
IF IEVS_type = "BEER" THEN CALL write_variable_in_CCOL_note("-----" & IEVS_year & " NON-WAGE MATCH(" & match_type & ") " & "(" & first_name & ") CLEARED CC-CLAIM ENTERED-----")
IF IEVS_type = "UBEN" THEN CALL write_variable_in_CCOL_note("-----" & IEVS_month & " NON-WAGE MATCH(" & match_type & ") " & "(" & first_name & ") CLEARED CC-CLAIM ENTERED-----")
IF IEVS_type = "UNVI" THEN CALL write_variable_in_CCOL_note("-----" & IEVS_year & " NON-WAGE MATCH(" & match_type & ") " & "(" & first_name & ") CLEARED CC-CLAIM ENTERED-----")
CALL write_bullet_and_variable_in_CCOL_note("Discovery date", discovery_date)
CALL write_bullet_and_variable_in_CCOL_note("Active Programs", programs)
CALL write_bullet_and_variable_in_CCOL_note("Source of income", income_source)
Call write_variable_in_CCOL_note("----- ----- ----- ----- -----")
IF OP_program <> "Select:" then Call write_variable_in_CCOL_note(OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number & " Amt $" & Claim_amount)
IF OP_program_II <> "Select:" then Call write_variable_in_CCOL_note(OP_program_II & " Overpayment " & OP_from_II & " through " & OP_to_II & " Claim # " & Claim_number_II & " Amt $" & Claim_amount_II)
IF OP_program_III <> "Select:" then Call write_variable_in_CCOL_note(OP_program_III & " Overpayment " & OP_from_III & " through " & OP_to_III & " Claim # " & Claim_number_III & " Amt $" & Claim_amount_III)
IF OP_program_IV <> "Select:" then Call write_variable_in_CCOL_note(OP_program_IV & " Overpayment " & OP_from_IV & " through " & OP_to_IV & " Claim # " & Claim_number_IV & " Amt $" & Claim_amount_IV)
IF HC_claim_number <> "" THEN
Call write_variable_in_CCOL_note("HC OVERPAYMENT " & HC_from & " through " & HC_to & " Claim #" & HC_claim_number & " Amt $" & HC_Claim_amount)
	Call write_bullet_and_variable_in_CCOL_note("Health Care responsible members", HC_resp_memb)
	Call write_bullet_and_variable_in_CCOL_note("Total Federal Health Care amount", Fed_HC_AMT)
	Call write_variable_in_CCOL_note("* Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
END IF
IF EI_checkbox = CHECKED THEN CALL write_variable_in_CCOL_note("* Earned Income Disregard Allowed")
IF EI_checkbox = UNCHECKED THEN CALL write_variable_in_CCOL_note("* Earned Income Disregard Not Allowed")
CALL write_bullet_and_variable_in_CCOL_note("Fraud referral made", fraud_referral)
CALL write_bullet_and_variable_in_CCOL_note("Income verification received", EVF_used)
CALL write_bullet_and_variable_in_CCOL_note("Date verification received", income_rcvd_date)
CALL write_bullet_and_variable_in_CCOL_note("Reason for overpayment", Reason_OP)
CALL write_bullet_and_variable_in_CCOL_note("Other responsible member(s)", OT_resp_memb)
'IF ECF_checkbox = CHECKED THEN CALL write_variable_in_CCOL_note("* DHS 2776E - Agency Cash Error Overpayment Worksheet form completed in ECF")
CALL write_variable_in_CCOL_note("----- ----- ----- ----- ----- ----- -----")
CALL write_variable_in_CCOL_note("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1")
PF3'


'-----------------------------------------------------------------------------------------CASENOTE
start_a_blank_CASE_NOTE
IF IEVS_type = "WAGE" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH"  & "(" & first_name & ") CLEARED CC-CLAIM ENTERED-----")
IF IEVS_type = "BEER" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_year & " NON-WAGE MATCH(" & match_type & ") " & "(" & first_name & ") CLEARED CC-CLAIM ENTERED-----")
IF IEVS_type = "UBEN" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_month & " NON-WAGE MATCH(" & match_type & ") " & "(" & first_name & ") CLEARED CC-CLAIM ENTERED-----")
IF IEVS_type = "UNVI" THEN CALL write_variable_in_CASE_NOTE("-----" & IEVS_year & " NON-WAGE MATCH(" & match_type & ") " & "(" & first_name & ") CLEARED CC-CLAIM ENTERED-----")
CALL write_bullet_and_variable_in_CASE_NOTE("Discovery date", discovery_date)
CALL write_bullet_and_variable_in_CASE_NOTE("Period", IEVS_period)
CALL write_bullet_and_variable_in_CASE_NOTE("Active Programs", programs)
CALL write_bullet_and_variable_in_CASE_NOTE("Source of income", income_source)
Call write_variable_in_CASE_NOTE("----- ----- ----- ----- -----")
Call write_variable_in_CASE_NOTE(OP_program & " Overpayment " & OP_from & " through " & OP_to & " Claim # " & Claim_number & " Amt $" & Claim_amount)
IF OP_program_II <> "Select:" then Call write_variable_in_CASE_NOTE(OP_program_II & " Overpayment " & OP_from_II & " through " & OP_to_II & " Claim # " & Claim_number_II & " Amt $" & Claim_amount_II)
IF OP_program_III <> "Select:" then	Call write_variable_in_CASE_NOTE(OP_program_III & " Overpayment " & OP_from_III & " through " & OP_to_III & " Claim # " & Claim_number_III & " Amt $" & Claim_amount_III)
IF OP_program_IV <> "Select:" then Call write_variable_in_CASE_NOTE(OP_program_IV & " Overpayment " & OP_from_IV & " through " & OP_to_IV & " Claim # " & Claim_number_IV & " Amt $" & Claim_amount_IV)
IF HC_claim_number <> "" THEN
	Call write_variable_in_case_note("HC OVERPAYMENT " & HC_from & " through " & HC_to & " Claim #" & HC_claim_number & " Amt $" & HC_Claim_amount)
	Call write_bullet_and_variable_in_CASE_NOTE("Health Care responsible members", HC_resp_memb)
	Call write_bullet_and_variable_in_CASE_NOTE("Total Federal Health Care amount", Fed_HC_AMT)
	Call write_variable_in_CASE_NOTE("* Emailed HSPHD Accounts Receivable for the medical overpayment(s)")
END IF
IF EI_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Earned Income Disregard Allowed")
IF EI_checkbox = UNCHECKED THEN CALL write_variable_in_case_note("* Earned Income Disregard Not Allowed")
CALL write_bullet_and_variable_in_case_note("Fraud referral made", fraud_referral)
CALL write_bullet_and_variable_in_case_note("Income verification received", EVF_used)
CALL write_bullet_and_variable_in_case_note("Date verification received", income_rcvd_date)
CALL write_bullet_and_variable_in_case_note("Reason for overpayment", Reason_OP)
CALL write_bullet_and_variable_in_case_note("Other responsible member(s)", OT_resp_memb)
IF transfer_to_worker <> "" THEN CALL write_variable_in_CASE_NOTE ("* Case transferred to X127" & transfer_to_worker)
'IF ECF_checkbox = CHECKED THEN CALL write_variable_in_CASE_NOTE("* DHS 2776E – Agency Cash Error Overpayment Worksheet form completed in ECF")
CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- ----- ----- -----")
CALL write_variable_in_CASE_NOTE("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1")
PF3 'to save casenote'

IF HC_claim_number <> "" THEN
	EmWriteScreen "x", 5, 3
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
	CALL create_outlook_email("HSPH.FIN.Unit.AR.Spaulding@hennepin.us", "","Claims entered for #" &  MAXIS_case_number & " Member # " & memb_number & " Date Overpayment Created: " & discovery_date & "HC Claim # " & HC_claim_number, "CASE NOTE" & vbcr & message_array,"", False)
END IF

transfer_to_worker = trim(transfer_to_worker)               'formatting the information entered in the dialog
transfer_to_worker = Ucase(transfer_to_worker)
'IF a transfer is needed (by entry of a transfer_to_worker in the Action dialog) the script will transfer it here
tansfer_message = ""            'some defaults
transfer_case = False
action_completed = TRUE

If transfer_to_worker <> "" Then        'If a transfer_to_worker was entered - we are attempting the transfer
	transfer_case = True
	CALL navigate_to_MAXIS_screen ("SPEC", "XFER")         'go to SPEC/XFER
	EMWriteScreen "x", 7, 16                               'transfer within county option
	transmit
	PF9                                                    'putting the transfer in edit mode
	EMreadscreen servicing_worker, 3, 18, 65               'checking to see if the transfer_to_worker is the same as the current_worker (because then it won't transfer)
	servicing_worker = trim(servicing_worker)
	IF servicing_worker = transfer_to_worker THEN          'If they match, cancel the transfer and save the information about the 'failure'
		action_completed = False
        transfer_message = "This case is already in the requested worker's number."
		PF10 'backout
		PF3 'SPEC menu
		PF3 'SELF Menu'
	ELSE                                                   'otherwise we are going for the tranfer
	    EMWriteScreen transfer_to_worker, 18, 61 		   'entering the worker ifnormation
	    transmit                                           'saving - this should then take us to the transfer menu
        EMReadScreen panel_check, 4, 2, 55                 'reading to see if we made it to the right place
        If panel_check = "XWKR" Then
            action_completed = False                       'this is not the right place
            transfer_message = "Transfer of this case to " & transfer_to_worker & " has failed."
            PF10 'backout
            PF3 'SPEC menu
            PF3 'SELF Menu'
        Else                                               'if we are in the right place - read to see if the new worker is the transfer_to_worker
            EMReadScreen new_pw, 3, 21, 20
            If new_pw <> transfer_to_worker Then           'if it is not the transfer_tow_worker - the transfer failed.
                action_completed = False
                transfer_message = "Transfer of this case to " & transfer_to_worker & " has failed."
            End If
        End If
	END IF
END IF
'' TODO call claim_referral_tracking

script_end_procedure_with_error_report("Overpayment case note entered and copied to CCOL please review case note to ensure accuracy.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------12/29/2021
'--Tab orders reviewed & confirmed----------------------------------------------12/29/2021
'--Mandatory fields all present & Reviewed--------------------------------------12/29/2021
'--All variables in dialog match mandatory fields-------------------------------12/29/2021
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------12/29/2021
'--CASE:NOTE Header doesn't look funky------------------------------------------12/29/2021
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------12/29/2021
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------12/29/2021
'--MAXIS_background_check reviewed (if applicable)------------------------------12/29/2021
'--PRIV Case handling reviewed -------------------------------------------------12/29/2021
'--Out-of-County handling reviewed----------------------------------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------12/29/2021
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------12/29/2021
'--Incrementors reviewed (if necessary)-----------------------------------------N/A
'--Denomination reviewed -------------------------------------------------------N/A
'--Script name reviewed---------------------------------------------------------12/29/2021
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub taks are complete-----------------------------------------12/29/2021
'--comment Code-----------------------------------------------------------------12/29/2021
'--Update Changelog for release/update------------------------------------------12/29/2021
'--Remove testing message boxes-------------------------------------------------12/29/2021
'--Remove testing code/unnecessary code-----------------------------------------12/29/2021
'--Review/update SharePoint instructions----------------------------------------12/29/2021
'--Review Best Practices using BZS page ----------------------------------------N/A
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------12/29/2021
'--Complete misc. documentation (if applicable)---------------------------------N/A
'--Update project team/issue contact (if applicable)----------------------------12/29/2021
