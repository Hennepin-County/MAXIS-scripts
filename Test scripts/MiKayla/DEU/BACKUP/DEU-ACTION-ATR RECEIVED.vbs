'GATHERING STATS===========================================================================================
name_of_script = "DEU-ACTION-ATR RECEIVED.vbs"
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
'Example: CALL changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
CALL changelog_update("11/07/2017", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

EMConnect ""
CALL MAXIS_case_number_finder (MAXIS_case_number)
memb_number = "01"
date_recieved = date & ""

'-----------------------------------------------------------------------------------------Initial dialog and do...loop
BeginDialog ATR_action_dialog, 0, 0, 181, 240, "ATR Recieved"
  EditBox 55, 5, 55, 15, MAXIS_case_number
  EditBox 155, 5, 20, 15, MEMB_Number
  EditBox 55, 25, 55, 15, date_recieved
  DropListBox 85, 45, 55, 15, "Select One:"+chr(9)+"1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"YEAR", select_quarter
  DropListBox 85, 65, 55, 15, "Select One:"+chr(9)+"WAGE"+chr(9)+"NON-WAGE", IEVS_period
  DropListBox 85, 85, 90, 15, "Select One:"+chr(9)+"MAIL"+chr(9)+"FAX"+chr(9)+"RCVD VERIFICATION", ATR_sent
  DropListBox 85, 105, 90, 15, "Select One:"+chr(9)+"DELETED DISQ"+chr(9)+"PENDING VERF"+chr(9)+"N/A", DISQ_action
  EditBox 65, 130, 110, 15, source_name
  EditBox 65, 150, 110, 15, source_address
  EditBox 65, 170, 110, 15, source_phone
  EditBox 65, 190, 110, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 65, 220, 50, 15
    CancelButton 125, 220, 50, 15
  Text 5, 10, 50, 10, "Case Number: "
  Text 120, 10, 30, 10, "MEMB #"
  Text 5, 30, 50, 10, "Date recieved:"
  Text 5, 50, 75, 10, "Match Period (quarter)"
  Text 5, 70, 65, 10, "Wage or Non-Wage"
  Text 45, 90, 30, 10, "ATR sent"
  Text 5, 110, 75, 10, "DISQ panel addressed"
  Text 10, 135, 50, 10, "Source Name:"
  Text 30, 155, 30, 10, "Address:"
  Text 15, 175, 45, 10, "Fax or Phone:"
  Text 20, 195, 45, 10, "Other Notes:"
EndDialog



Do
	Do
        err_msg = "" 
		Dialog ATR_action_dialog
		IF ButtonPressed = 0 THEN StopScript
		IF IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
		IF select_quarter = "Select One:" THEN err_msg = err_msg & vbNewLine & "Please select a quarter for the match"
		IF match_type = "Select One:" THEN err_msg = err_msg & vbNewLine & "Please select a match type"
		IF ATR_sent = "Select One:" THEN err_msg = err_msg & vbNewLine & "Please select how ATR was sent"
		IF DISQ_action = "Select One:" THEN err_msg = err_msg & vbNewLine & "Please advise IF DISQ was updated"
		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine	
    Loop until err_msg = ""	
 	CALL check_for_password(are_we_passworded_out)
LOOP UNTIL check_for_password(are_we_passworded_out) = False	

'-------------------------------------------------------------------------------------------Defaulting the quarters 
IF select_quarter = "1" THEN
                IEVS_period = "01-" & CM_yr & "/03-" & CM_yr
ELSEIF select_quarter = "2" THEN
                IEVS_period = "04-" & CM_yr & "/06-" & CM_yr
ELSEIF select_quarter = "3" THEN
                IEVS_period = "07-" & CM_yr & "/09-" & CM_yr
ELSEIF select_quarter = "4" THEN
                IEVS_period = "10-" & CM_minus_6_yr & "/12-" & CM_minus_6_yr
ELSEIF select_quarter = "YEAR" THEN
				IEVS_period = right(DatePart("yyyy",DateAdd("yyyy", -1, date)), 2) 
END IF

msgbox IEVS_period

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


Row = 7	
DO 
	EMReadScreen IEVS_match, 11, row, 47 
	IF trim(IEVS_match) = "" THEN script_end_procedure("IEVS match for the selected period could not be found. The script will now end.")
	IF IEVS_match = IEVS_period THEN
		msgbox IEVS_period
		EXIT DO
	ELSE 
		row = row + 1
		msgbox "row: " & row 
		IF row = 17 THEN 
			PF8
			row = 7
		END IF
	END IF
Loop until IEVS_period = select_quarter 

EMReadScreen multiple_match, 11, row + 1, 47 
IF multiple_match = IEVS_period THEN 
	msgbox("More than one match exists for this time period. Determine the match you'd like to clear, and put your cursor in front of that match." & vbcr & "Press OK once match is determined.")
	EMSendKey "U"
	transmit
ELSE 
	CALL write_value_and_transmit("U", row, 3)   'navigates to IULA
END IF 

'----------------------------------------------------------------------------------------------------IULA
'Entering the IEVS match & reading the dIFference notice to ensure this has been sent
'Reading potential errors for out-of-county cases
EMReadScreen OutOfCounty_error, 12, 24, 2
IF OutOfCounty_error = "MATCH IS NOT" THEN
	script_end_procedure("Out-of-county case. Cannot update.")
ELSE
	IF IEVS_type = "WAGE" THEN
		EMReadScreen quarter, 1, 8, 14
		EMReadScreen IEVS_year, 4, 8, 22
		IF quarter <> select_quarter THEN script_end_procedure("Match period does not match the selected match period. The script will now end.")
	ELSEIF IEVS_type = "NON-WAGE" THEN
		EMReadScreen Nonwage_year , 2, 8, 15
		Nonwage_year = "20" & Nonwage_year
	END IF
END IF 

'----------------------------------------------------------------------------------------------------Client name
EMReadScreen client_name, 35, 5, 24
'Formatting the client name for the spreadsheet
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
IF instr(first_name, " ") THEN   						'IF there is a middle initial in the first name, THEN it removes it
	length = len(first_name)                        	'trimming the 1st name
	position = InStr(first_name, " ")               	'establishing the length of the variable
	first_name = Left(first_name, position-1)       	'trims the middle initial off of the first name
END IF

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
EMReadScreen employer_info, 27, 8, 37
employer_info = trim(employer_info)

IF instr(employer_info, "AMT: $") THEN 
    length = len(employer_info) 						  'establishing the length of the variable
    position = InStr(employer_info, "AMT: $")    		      'sets the position at the deliminator  
    employer_name = Left(employer_info, position)  'establishes employer as being before the deliminator
ELSE 
    employer_name = employer_info
END IF 

EMReadScreen dIFf_notice, 1, 14, 37
EMReadScreen dIFf_date, 10, 14, 68
dIFf_date = trim(dIFf_date)
IF dIFf_date <> "" THEN dIFf_date = replace(dIFf_date, " ", "/")


'--------------------------------------------------------------------sending the notice in IULA
EMwritescreen "005", 12, 46 'writing the resolve time to read for later
EMwritescreen "Y", 15, 37 'send Notice
transmit 'this will take us to IULB'\
ROW = 8
EMReadScreen IULB_first_line, 1, row, 6
IF IULB_first_line = "" THEN 
	EMwritescreen "ATR RECEIVED " & date_recieved, row, 6
ELSE 
	ROW = 9
	CALL clear_line_of_text(row, 6)
	EMwritescreen "ATR RECEIVED " & date_recieved, row, 6	
END IF 		

msgbox "responded to dIFference notice updated"
transmit'exiting IULA, helps prevent errors when going to the case note

'--------------------------------------------------------------------The case note & case note related code

Due_date = dateadd("d", 10, date)	'defaults the due date for all verIFications at 10 days

'Updated IEVS_period to write into case note
IF select_quarter = "1" THEN IEVS_quarter = "1ST"
IF select_quarter = "2" THEN IEVS_quarter = "2ND"
IF select_quarter = "3" THEN IEVS_quarter = "3RD"
IF select_quarter = "4" THEN IEVS_quarter = "4TH"
IF select_quarter = "YEAR" THEN IEVS_quarter = Nonwage_year

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
 
Due_date = dateadd("d", 10, date)	'defaults the due date for all verIFications at 10 days
IEVS_period = replace(IEVS_period, "/", " to ")
dIFf_date = replace(dIFf_date, " ", "/")

start_a_blank_CASE_NOTE
	IF IEVS_quarter <> "YEAR" THEN 
		CALL write_variable_in_CASE_NOTE ("-----" & IEVS_quarter & " QTR " & IEVS_year & "WAGE MATCH (" & first_name & ") ATR RECIEVED-----")
	ELSE 
		CALL write_variable_in_CASE_NOTE ("-----" & IEVS_year & " WAGE MATCH (" & first_name & ") ATR RECIEVED-----")
	END IF	
	CALL write_bullet_and_variable_in_CASE_NOTE("Period", IEVS_period)
	CALL write_bullet_and_variable_in_CASE_NOTE("Active Programs", programs)
	CALL write_variable_in_CASE_NOTE("Source information: " & source_name & "  " & source_address)
	CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- ----- ----- -----")
	CALL write_variable_in_CASE_NOTE("* Date ATR received: " & date_recieved)
	IF DISQ_action = "DELETED DISQ" THEN CALL write_variable_in_CASE_NOTE("* Updated DISQ panel")
	IF DISQ_action = "PENDING VERF" THEN CALL write_variable_in_CASE_NOTE("* Pending verification of income or asset")
	IF ATR_sent <> "RCVD VERIFICATION" THEN 
		CALL write_variable_in_CASE_NOTE("* Sent via: " & ATR_sent & ", " & source_phone)
		CALL write_bullet_and_variable_in_case_note("Due Date", Due_date)
		CALL write_variable_in_CASE_NOTE("---DEU WILL PROCESS WHEN EMPLOYMENT VERIFICATION IS RETURNED. TEAM CAN REINSTATE CASE IF ALL NECESSARY PAPERWORK TO REINSTATE HAS BEEN RECIEVED---")
    ELSE
		CALL write_variable_in_CASE_NOTE("---TEAM CAN REIN CASE IF ALL NECESSARY PAPERWORK TO REINS HAS BEEN RCVD---")
	END IF
	CALL write_bullet_and_variable_in_case_note("Other Notes", other_notes)
	CALL write_variable_in_CASE_NOTE("* IEVP updated as responded to dIFference notice - YES ")
	CALL write_variable_in_CASE_NOTE ("----- ----- ----- ----- ----- ----- -----")
	CALL write_variable_in_CASE_NOTE ("DEBT ESTABLISHMENT UNIT 612-348-4290 EXT 1-1-1")

script_end_procedure("ATR case note updated successfully." & vbNewLine & "Please remember to update/delete the DISQ panel")


 