'GATHERING STATS===========================================================================================
name_of_script = "DEU-MATCH CLEARED.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 150
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
call changelog_update("07/12/2017", "Updated.", "MiKayla Handley, Hennepin County")
call changelog_update("05/15/2017", "Initial version.", "Ilse Ferris, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to MAXIS & grabs case number
EMConnect ""

'CHECKS TO MAKE SURE THE WORKER IS ON THEIR DAIL
EMReadscreen dail_check, 4, 2, 48
If dail_check <> "DAIL" then script_end_procedure("You are not in your dail. This script will stop.")

'----------------------------------------------------------------------------------------------------DAIL
'TYPES A "T" TO BRING THE SELECTED MESSAGE TO THE TOP
EMSendKey "t"
'checking for an active MAXIS session
Call check_for_MAXIS(False)

EMReadScreen IEVS_type, 4, 6, 6 'read the DAIL msg'
If IEVS_type <> "WAGE" then 
	if IEVS_type <> "BEER" then 
		script_end_procedure("This is not a IEVS match. Please select a IEVS match DAIL, and run the script again.")
	End if 
End if 

EMReadScreen MAXIS_case_number, 8, 5, 73
MAXIS_case_number= TRIM(MAXIS_case_number)

'----------------------------------------------------------------------------------------------------IEVS
'Navigating deeper into the match interface
CALL write_value_and_transmit("I", 6, 3)   		'navigates to INFC 
CALL write_value_and_transmit("IEVP", 20, 71)   'navigates to IEVP
EMReadScreen error_msg, 7, 24, 2
If error_msg = "NO IEVS" then script_end_procedure("An error occured in IEVP, please process manually.")'checking for error msg'

row = 7
'Ensuring that match has not already been resolved.
Do
	EMReadScreen days_pending, 5, row, 72
	days_pending = trim(days_pending)
	If IsNumeric(days_pending) = false then 
		script_end_procedure("No pending IEVS match found. Please review IEVP.")
	ELSE
		'Entering the IEVS match & reading the difference notice to ensure this has been sent
		EMReadScreen IEVS_period, 11, row, 47
		EMReadScreen start_month, 2, row, 47
		EMReadScreen end_month, 2, row, 53
		If trim(start_month) = "" or trim(end_month) = "" then 
			Found_match = False
		else
			month_difference = abs(end_month) - abs(start_month)
			If (IEVS_type = "WAGE" and month_difference = 2) then 'ensuring if it is a wage the match is a quater'
				found_match = true
				exit do
			Elseif (IEVS_type = "BEER" and month_difference = 11) then  'ensuring that if it a beer that the match is a year'
				found_match = True
				exit do
			End if
		End if
		row = row + 1
	END IF
Loop until row = 17

If found_match = False then script_end_procedure("No pending IEVS match found. Please review IEVP.")

'----------------------------------------------------------------------------------------------------IULA
'Entering the IEVS match & reading the difference notice to ensure this has been sent
CALL write_value_and_transmit("U", row, 3)   'navigates to IULA
'Reading potential errors for out-of-county cases
EMReadScreen OutOfCounty_error, 12, 24, 2
IF OutOfCounty_error = "MATCH IS NOT" then
	script_end_procedure("Out-of-county case. Cannot update.")
else
	IF IEVS_type = "WAGE" then
		EMReadScreen quarter, 1, 8, 14
		EMReadScreen IEVS_year, 4, 8, 22
	Elseif IEVS_type = "BEER" then
		EMReadScreen IEVS_year, 2, 8, 15
		IEVS_year = "20" & IEVS_year
	End if
End if 

'----------------------------------------------------------------------------------------------------Client name
EMReadScreen client_name, 35, 5, 24
'Formatting the client name for the spreadsheet
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
Active_Programs =trim(Active_Programs)

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

'----------------------------------------------------------------------------------------------------Employer info & differnce notice info
EMReadScreen employer_info, 27, 8, 37
employer_info = trim(employer_info)

If instr(employer_info, " AMT: $") then 
    length = len(employer_info) 						  'establishing the length of the variable
    position = InStr(employer_info, " AMT: $")    		      'sets the position at the deliminator  
    employer_info = Left(employer_info, position)  'establishes employer as being before the deliminator
Else 
    employer_info = employer_info
End if 

EMReadScreen diff_notice, 1, 14, 37
EMReadScreen diff_date, 10, 14, 68
diff_date = trim(diff_date)
If diff_date <> "" then diff_date = replace(diff_date, " ", "/")

'----------------------------------------------------------------------------------------------------initial case number dialog
BeginDialog cleared_match_dialog, 0, 0, 331, 145, "IEVS Match Cleared"
  Text 10, 20, 130, 10, "Case number: " & MAXIS_case_number
  Text 140, 20, 185, 10, "Client name: " & client_name
  Text 10, 40, 125, 10, "Open prog(s): " & programs
  Text 140, 40, 185, 10, "Income source: " & employer_info
  Text 10, 60, 125, 10, "Difference notice sent?:  "& diff_notice
  Text 140, 60, 120, 10, "Date sent: "& diff_date
  EditBox 100, 85, 30, 15, resolve_time
  DropListBox 205, 85, 120, 15, "Select one..."+chr(9)+"BC - Case Closed"+chr(9)+"BN - Already known, No Savings"+chr(9)+"BE - Child"+chr(9)+"BE - No Change"+chr(9)+"CC - Claim Entered", Cleared_status
  DropListBox 245, 105, 80, 15, "Select one..."+chr(9)+"Yes"+chr(9)+"No", change_response_to_notice
  EditBox 55, 125, 170, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 230, 125, 45, 15
    CancelButton 280, 125, 45, 15
  Text 10, 130, 45, 10, "Other notes:"
  Text 15, 90, 85, 10, "Resolve time (in minutes):"
  Text 15, 110, 230, 10, "If 'Responded to the difference notice' needs to be updated, select one:"
  GroupBox 5, 5, 325, 75, "IEVS_type &  MATCH"
  Text 135, 90, 70, 10, "Select resolved type:"
EndDialog

Do 
	Do 
		err_msg = ""
	    dialog cleared_match_dialog
        cancel_confirmation
        If IsNumeric(resolve_time) = false or len(resolve_time) > 3 then err_msg = err_msg & vbNewLine & "* Enter a valid numeric resolved time."
		If Cleared_status = "Select one..." then err_msg = err_msg & vbNewLine & "* Enter an resolved option."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine		
	Loop until err_msg = ""
	
	'CHECKING FOR MAXIS WITHOUT TRANSMITTING SINCE THIS WILL NAVIGATE US AWAY FROM THE AREA WE ARE AT
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

'Updating the 'Responded to the difference notice' field if selected by the user in the dialog
If change_response_to_notice = "Yes" then EMWriteScreen "Y", 15, 37
If change_response_to_notice = "No"  then EMWriteScreen "N", 15, 37

'clearing all programs on IULA 
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
	EMWriteScreen Cleared_status, row + 1, col + 1
Next 

CALL write_value_and_transmit(resolve_time, 12, 46) 
'resolved notes depending on the Cleared_status
If Cleared_status = "BC - Case Closed" 	then EMWriteScreen "Case closed. " & other_notes, 8, 6   							'BC
If Cleared_status = "BE - No Change" then EMWriteScreen "No change. " & other_notes, 8, 6 									'BE
If cleared_status = "BE - Child" then EMWriteScreen "No change, minor child income excluded. " & other_notes, 8, 6 			'BE - child
If Cleared_status = "BN - Already known, No Savings" then EMWriteScreen "Already known - No savings. " & other_notes, 8, 6 	'BN
If Cleared_status = "CC - Claim Entered" then EMWriteScreen "Claim entered. " & other_notes, 8, 6 						 	'CC
transmit				

'back on the IEVP menu, making sure that the match cleared
EMReadScreen days_pending, 5, 7, 72
days_pending = trim(days_pending)
If IsNumeric(days_pending) = true then 
	match_cleared = False 
	script_end_procedure("This match did not appear to clear. Please check case, and try again.")
Else 
	match_cleared = true
End if 

If match_cleared = true then 
'Formatting for the case note----------------------------------------------------------------------------------------------------
    If IEVS_type = "WAGE" then
    	'Updated IEVS_period to write into case note
    	If quarter = 1 then IEVS_quarter = "1ST"
    	If quarter = 2 then IEVS_quarter = "2ND"
    	If quarter = 3 then IEVS_quarter = "3RD"
    	If quarter = 4 then IEVS_quarter = "4TH"
    End if
     
    'adding specific wording for case note header for each cleared status
    If Cleared_status = "BC - Case Closed" then cleared_header_info = " (" & first_name & ") CLEARED BC-CASE CLOSED"
    If Cleared_status = "BE - No Change" then cleared_header_info = " (" & first_name & ") CLEARED BE-NO CHANGE"
	If cleared_status = "BE - Child" then cleared_header_info = " (" & first_name & ") CLEARED BE-NO CHANGE- MINOR CHILD"
    If Cleared_status = "BN - Already known, No Savings" then cleared_header_info = " (" & first_name & ") CLEARED BN-KNOWN"
    If Cleared_status = "CC - Claim Entered" then cleared_header_info = " (" & first_name & ") CLEARED CC-CLAIM ENTERED"
    
    IEVS_period = replace(IEVS_period, "/", " to ")
    'The casenote----------------------------------------------------------------------------------------------------
    start_a_blank_CASE_NOTE
    If IEVS_type = "WAGE" then Call write_variable_in_CASE_NOTE("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE INCOME" & cleared_header_info & "-----")
    If IEVS_type = "BEER" then Call write_variable_in_CASE_NOTE("-----" & IEVS_year & " NON WAGE INCOME(B)" & cleared_header_info & "-----")
    Call write_bullet_and_variable_in_CASE_NOTE("Period", IEVS_period)
    Call write_bullet_and_variable_in_CASE_NOTE("Programs open", programs)
    Call write_bullet_and_variable_in_CASE_NOTE("Employer name", employer_info)
    Call write_variable_in_CASE_NOTE ("----- ----- -----")
    If Cleared_status = "BN - Already known, No Savings" or Cleared_status = "BE - No Change" then Call write_variable_in_CASE_NOTE("CLIENT REPORTED EARNINGS. INCOME IS IN STAT PANELS AND BUDGETED.")
    If cleared_status = "BE - Child" then Call write_variable_in_CASE_NOTE("INCOME IS EXCLUDED FOR MINOR CHILD IN SCHOOL.")
	If Cleared_status <> "CC - Claim Entered" then Call write_variable_in_CASE_NOTE("NO OVERPAYMENTS OR SAVINGS RELATED TO THIS MATCH.")
    Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
    Call write_variable_in_CASE_NOTE("----- ----- ----- ----- ----- ----- -----")
    Call write_variable_in_CASE_NOTE("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1")
End if 

script_end_procedure("")