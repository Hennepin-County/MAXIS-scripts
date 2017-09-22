'GATHERING STATS===========================================================================================
name_of_script = "DEU-DISQ DELETED.vbs"
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
call changelog_update("07/13/2017", "Updated title and removed request line.", "MiKayla Handley, Hennepin County")
call changelog_update("06/06/2017", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

EMConnect ""
CALL MAXIS_case_number_finder (MAXIS_case_number)
memb_number = "01"
DISQ_deleted_Date = date & ""

'-----------------------------------------------------------------------------------------Initial dialog and do...loop
BeginDialog DISQ_deleted, 0, 0, 186, 120, "DISQ Deleted"
  EditBox 55, 20, 55, 15, MAXIS_case_number
  EditBox 155, 20, 20, 15, MEMB_Number
  DropListBox 90, 40, 55, 15, "Select One..."+chr(9)+"1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4", select_quarter
  EditBox 90, 60, 55, 15, DISQ_deleted_Date
  EditBox 50, 80, 125, 15, Other_Notes
  ButtonGroup ButtonPressed
    OkButton 70, 100, 50, 15
    CancelButton 125, 100, 50, 15
  Text 120, 25, 30, 10, "MEMB #"
  Text 5, 5, 180, 10, "**STAT/DISQ panel will NOT be updated by the script**"
  Text 5, 85, 45, 10, "Other Notes:"
  Text 5, 25, 50, 10, "Case Number: "
  Text 15, 65, 45, 10, "Date deleted:"
  Text 15, 45, 70, 10, "IEVS Match Quarter"
  Text 5, 5, 180, 10, "**STAT/DISQ panel will NOT be deleted by the script**"
EndDialog

Do
	Do
        err_msg = "" 
		Dialog
		IF ButtonPressed = 0 then StopScript
		IF IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine	
    Loop until err_msg = ""	
 	Call check_for_password(are_we_passworded_out)
LOOP UNTIL check_for_password(are_we_passworded_out) = False	

'----------------------------------------------------------------------------------------------------Creating the quarter
CM_minus_6_yr =  right(DatePart("yyyy", DateAdd("m", -6, date)), 2)

''Fun with dates! Defalting the quarters 
IF select_quarter = "1" THEN
                IEVS_period = "01-" & CM_yr & "/03-" & CM_yr
ELSEIF select_quarter = "2" THEN
                IEVS_period = "04-" & CM_yr & "/06-" & CM_yr
ELSEIF select_quarter = "3" THEN
                IEVS_period = "07-" & CM_yr & "/09-" & CM_yr
ELSEIF select_quarter = "4" THEN
                IEVS_period = "10-" & CM_minus_6_yr & "/12-" & CM_minus_6_yr
END IF

'----------------------------------------------------------------------------------------------------IEVS

Call navigate_to_MAXIS_screen("STAT", "MEMB")
EMwritescreen memb_number, 20, 76
transmit

EMReadscreen SSN_number_read, 11, 7, 42
SSN_number_read = replace(SSN_number_read, " ", "") 

Call navigate_to_MAXIS_screen("INFC" , "____")  
CALL write_value_and_transmit("IEVP", 20, 71) 
CALL write_value_and_transmit(SSN_number_read, 3, 63) '

EMReadScreen edit_error, 2, 24, 2
edit_error = trim(edit_error)
IF edit_error <> "" then script_end_procedure("No IEVS matches and/ or could not access IEVP.")

Row = 7	
Do 
	EMReadScreen IEVS_match, 11, row, 47 
	If trim(IEVS_match) = "" THEN script_end_procedure("IEVS match for the selected period could not be found. The script will now end.")
	If IEVS_match = IEVS_period then 
		'msgbox "Found it!"
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
	msgbox("More than one match exists for this time period. Determine the match you'd like to clear, and put your cursor in front of that match." & vbcr & "Press OK once match is determined.")
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
else
	IF IEVS_type = "WAGE" then
		EMReadScreen quarter, 1, 8, 14
		EMReadScreen IEVS_year, 4, 8, 22
		If quarter <> select_quarter then script_end_procedure("Match period does not match the selected match period. The script will now end.")
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

If instr(employer_info, "AMT: $") then 
    length = len(employer_info) 						  'establishing the length of the variable
    position = InStr(employer_info, "AMT: $")    		      'sets the position at the deliminator  
    employer_name = Left(employer_info, position)  'establishes employer as being before the deliminator
Else 
    employer_name = employer_info
End if 

EMReadScreen diff_notice, 1, 14, 37
EMReadScreen diff_date, 10, 14, 68
diff_date = trim(diff_date)
If diff_date <> "" then diff_date = replace(diff_date, " ", "/")

PF3		'exiting IULA, helps prevent errors when going to the case note

'-----------------------------------------------------------------------------------'for the case notes
'Updated IEVS_period to write into case note
If select_quarter = "1" then IEVS_Quarter = "1ST"
If select_quarter = "2" then IEVS_Quarter = "2ND"
If select_quarter = "3" then IEVS_Quarter = "3RD"
If select_quarter = "4" then IEVS_Quarter = "4TH"

programs = ""
IF instr(Active_Programs, "D") then programs = programs & "DWP, "
IF instr(Active_Programs, "F") then programs = programs & "Food Support, "
IF instr(Active_Programs, "H") then programs = programs & "Health Care, "
IF instr(Active_Programs, "M") then programs = programs & "Medical Assistance, "
IF instr(Active_Programs, "S") then programs = programs & "MFIP, "
'trims excess spaces of programs 
programs = trim(programs)
'takes the last comma off of programs when autofilled into dialog
IF right(programs, 1) = "," THEN programs = left(programs, len(programs) - 1) 
 
Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days
IEVS_period = replace(IEVS_period, "/", " to ")
diff_date = replace(diff_date, " ", "/")

start_a_blank_CASE_NOTE
	Call write_variable_in_CASE_NOTE ("-----" & IEVS_Quarter & " QTR " & IEVS_year & " WAGE MATCH (" & first_name & ") DISQ DELETED-----")
	Call write_bullet_and_variable_in_CASE_NOTE("Period", IEVS_period)
	Call write_bullet_and_variable_in_CASE_NOTE("Active Programs", programs)
	Call write_bullet_and_variable_in_CASE_NOTE("Employer info", Employer_info)
	Call write_variable_in_CASE_NOTE("----- ----- ----- ----- ----- ----- -----")
	Call write_variable_in_CASE_NOTE("* DEU STAT/DISQ CODE 14 (WAGE) HAS BEEN DELETED")
	Call write_variable_in_CASE_NOTE("* Date ATR received: " & date_atr_rcvd)
	Call write_variable_in_CASE_NOTE("* Employer: " & employer_info)
	Call write_variable_in_CASE_NOTE("---" & "DEU WILL PROCESS WHEN EMPLOYMENT VERIFICATION IS RETURNED. TEAM CAN REINSTATE CASE IF ALL NECESSARY PAPERWORK TO REINSTATE HAS BEEN RECIEVED---")
	Call write_bullet_and_variable_in_case_note("Other Notes", Other_Notes)
	Call write_variable_in_CASE_NOTE ("----- ----- ----- ----- ----- ----- -----")
	Call write_variable_in_CASE_NOTE ("DEBT ESTABLISHMENT UNIT 612-348-4290 EXT 1-1-1")

script_end_procedure("")


 