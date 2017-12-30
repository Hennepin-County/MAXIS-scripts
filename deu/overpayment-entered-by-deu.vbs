'GATHERING STATS===========================================================================================
name_of_script = "OVERPAYMENT CLAIM ENTERED.vbs"
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

BeginDialog OP_Cleared_dialog, 0, 0, 281, 265, "Overpayment Claim Entered"
  EditBox 55, 5, 35, 15, MAXIS_case_number
  DropListBox 140, 5, 55, 15, "Select One:"+chr(9)+"1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"YEAR"+chr(9)+"LAST YEAR"+chr(9)+"OTHER", select_quarter
  EditBox 240, 5, 35, 15, memb_number
  DropListBox 55, 25, 35, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", fraud_referral
  DropListBox 140, 25, 55, 15, "Select One:"+chr(9)+"WAGE"+chr(9)+"BEER"+chr(9)+"UNVI", IEVS_type
  DropListBox 240, 25, 35, 15, "Select One:"+chr(9)+"YES"+chr(9)+"NO", collectible_dropdown
  EditBox 35, 55, 35, 15, OP_1
  EditBox 90, 55, 35, 15, OP_to_1
  EditBox 160, 55, 35, 15, Claim_1
  EditBox 220, 55, 45, 15, AMT_1
  EditBox 35, 75, 35, 15, OP_2
  EditBox 90, 75, 35, 15, OP_to_2
  EditBox 160, 75, 35, 15, Claim_2
  EditBox 220, 75, 45, 15, Amt_2
  EditBox 35, 95, 35, 15, OP_3
  EditBox 90, 95, 35, 15, OP_to_3
  EditBox 160, 95, 35, 15, Claim_3
  EditBox 220, 95, 45, 15, AMT_3
  EditBox 35, 115, 35, 15, OP_4
  EditBox 90, 115, 35, 15, OP_to_4
  EditBox 160, 115, 35, 15, Claim_4
  EditBox 220, 115, 45, 15, AMT_4
  EditBox 75, 140, 80, 15, collectible_reason
  EditBox 75, 160, 80, 15, EVF_date
  EditBox 230, 140, 45, 15, OT_resp_memb
  EditBox 230, 160, 45, 15, Fed_HC_AMT
  EditBox 230, 180, 45, 15, HC_resp_memb
  CheckBox 5, 205, 120, 10, "Earned Income disregard allowed", EI_checkbox
  EditBox 60, 220, 215, 15, Reason_OP
  EditBox 75, 180, 80, 15, other_programs
  CheckBox 5, 245, 160, 10, "Check here if there is no wage match to act on", no_match_action
  ButtonGroup ButtonPressed
    OkButton 180, 245, 45, 15
    CancelButton 230, 245, 45, 15
  Text 5, 10, 50, 10, "Case Number: "
  Text 95, 10, 45, 10, "Match Period:"
  Text 205, 10, 30, 10, "MEMB #:"
  Text 5, 30, 50, 10, "Fraud referral:"
  Text 95, 30, 40, 10, "IEVS Type:"
  Text 200, 30, 40, 10, "Collectible?"
  GroupBox 5, 45, 270, 90, "Overpayment Information"
  Text 10, 60, 25, 10, "From:"
  Text 75, 60, 10, 10, "To:"
  Text 130, 60, 25, 10, "Claim #"
  Text 200, 60, 20, 10, "AMT:"
  Text 10, 80, 20, 10, "From:"
  Text 75, 80, 10, 10, "To:"
  Text 130, 80, 25, 10, "Claim #"
  Text 200, 80, 20, 10, "AMT:"
  Text 10, 100, 20, 10, "From:"
  Text 75, 100, 10, 10, "To:"
  Text 130, 100, 25, 10, "Claim #"
  Text 200, 100, 20, 10, "AMT:"
  Text 10, 120, 20, 10, "From:"
  Text 75, 120, 10, 10, "To:"
  Text 200, 120, 20, 10, "AMT:"
  Text 130, 120, 25, 10, "Claim #"
  Text 5, 145, 65, 10, "Collectible Reason:"
  Text 5, 165, 60, 10, "Income verif used:"
  Text 160, 145, 65, 10, "HC resp. members:"
  Text 160, 165, 65, 10, "Total FED HC AMT:"
  Text 160, 185, 60, 10, "Other resp. memb:"
  Text 5, 225, 50, 10, "Reason for OP: "
  Text 10, 185, 55, 10, "Other programs:"
EndDialog


Do
	err_msg = ""
	dialog OP_Cleared_dialog
	IF buttonpressed = 0 then stopscript 
	IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
	IF select_quarter = "Select One:" THEN err_msg = err_msg & vbnewline & "* You must select a match period entry."
	IF IEVS_type = "Select One:" THEN err_msg = err_msg & vbnewline & "* You must select a match type entry."
    IF OP_1 = "" THEN err_msg = err_msg & vbnewline & "* You must have an overpayment entry."
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
CALL write_bullet_and_variable_in_case_note("Income verification received", EVF_date)
CALL write_bullet_and_variable_in_case_note("Other responsible member(s)", OT_resp_memb)
CALL write_bullet_and_variable_in_case_note("Reason for overpayment", Reason_OP) 
CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- -----")
CALL write_variable_in_CASE_NOTE("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1") 
IF instr(Active_Programs, "HC") THEN CALL create_outlook_email("", "mikayla.handley@hennepin.us", "Claims entered for #" &  MAXIS_case_number, "Member #: " & memb_number & vbcr & "Date Overpayment Created: " & OP_Date & vbcr & "Programs: " & program_droplist & vbcr & "See case notes for further details.", "", False)