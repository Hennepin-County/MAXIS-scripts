'GATHERING STATS===========================================================================================
name_of_script = "DEU-SEND DIFF MATCH NOTICE.vbs"
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
call changelog_update("08/17/2017", "Updated functionality when searching for more than one match. Also added information to closing message box if a wage match needing a difference notice cannot be found.", "Ilse Ferris, Hennepin County")
call changelog_update("05/17/2017", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

EMConnect ""
CALL check_for_MAXIS(False)

'CHECKS TO MAKE SURE THE WORKER IS ON THEIR DAIL
EMReadscreen dail_check, 4, 2, 48
If dail_check <> "DAIL" then script_end_procedure("You are not in your dail. This script will stop.")

'TYPES A "T" TO BRING THE SELECTED MESSAGE TO THE TOP
EMSendKey "t"
transmit

EMReadScreen IEVS_type, 4, 6, 6 'read the DAIL msg'
If IEVS_type <> "WAGE" then 
	if IEVS_type <> "BEER" then 
		script_end_procedure("This is not a IEVS match. Please select a non-wage match DAIL, and run the script again.")
	End if 
End if 

EMReadScreen MAXIS_case_number, 8, 5, 73
MAXIS_case_number= TRIM(MAXIS_case_number)

''----------------------------------------------------------------------------------------------------IEVS
'Navigating deeper into the match interface
CALL write_value_and_transmit("I", 6, 3)   		'navigates to INFC 
CALL write_value_and_transmit("IEVP", 20, 71)   'navigates to IEVP
EMReadScreen error_msg, 7, 24, 2
If error_msg = "NO IEVS" then script_end_procedure("An error occured in IEVP, please process manually.")'checking for error msg'

'----------------------------------------------------------------------------------------------------THE DIALOG 
BeginDialog SEND_WAGE_MATCH_DIFF_NOTICE_dialog , 0, 0, 131, 85, "ACTION-SEND WAGE MATCH DIFF NOTICE"
  EditBox 65, 5, 60, 15, maxis_case_number
  EditBox 65, 25, 60, 15, ATR_on_file_date
  CheckBox 5, 45, 110, 15, "Update Claim Referral Tracking (CASH/FS ONLY)", Claim_Referral_checkbox
  ButtonGroup ButtonPressed
    OkButton 20, 65, 50, 15
    CancelButton 75, 65, 50, 15
  Text 5, 30, 55, 10, "ATR on file Date:"
  Text 5, 10, 45, 10, "Case number:"
EndDialog
Do
	Do
		err_msg = ""
		Dialog SEND_WAGE_MATCH_DIFF_NOTICE_dialog
		If ButtonPressed = 0 then StopScript
		IF len(MAXIS_case_number) > 8 or isnumeric(MAXIS_case_number) = false THEN err_msg = err_msg & vbCr & "You must enter a valid case number."		'mandatory field 
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	
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

'----------------------------------------------------------------------------------------------------IEVS AGAIN
'Do loop established to handle more than one match that has not been resolved yet. User has the option to select another match if this is not the correct match OR a notice was already sent on one match, but not the other. 
match_num = 0	'establishing the count of pending WAGE matches we'll find
diff_sent = 0	'establishing the count of difference notices sent we'll find
row = 7
Do 
    'Ensuring that match has not already been resolved.
    Do
    	EMReadScreen days_pending, 5, row, 72
    	days_pending = trim(days_pending)
    	If IsNumeric(days_pending) = true then
            'Entering the IEVS match & reading the difference notice to ensure this has been sent
        	EMReadScreen IEVS_period, 11, row, 47
    		EMReadScreen start_month, 2, row, 47
    		EMReadScreen end_month, 2, row, 53
    		
    		If trim(start_month) = "" then start_month = "0"
    		If trim(end_month) = "" then end_month = "0"
    		month_difference = abs(end_month) - abs(start_month)
    		If month_difference = 2 then 'ensuring if it is a wage the match is a quater'
				match_num = match_num + 1
    			access_match = True
    			exit do
    		else
    			row = row + 1
    		End if
		Else 
			row = row + 1
    	END IF
    Loop until row = 17
	If row = 17 then script_end_procedure("Could not find pending WAGE match that requires a difference notice sent. Please review the case, and try again." & vbcr & vbcr & "Number of wage matches found: " & match_num & vbcr & "Number of difference notices sent: " & diff_sent)
	
	If access_match <> True then script_end_procedure("This WAGE match is not for a quarter. Please process manually.")
	
	CALL write_value_and_transmit("U", row, 3)   'navigates to IULA
	'Reading potential errors for out-of-county cases
	EMReadScreen OutOfCounty_error, 12, 24, 2
	IF OutOfCounty_error = "MATCH IS NOT" then script_end_procedure("Out-of-county case. Cannot update.")
	
	'sending the notice
	EMReadScreen IULA, 4, 2, 52
	IF IULA <> "IULA" then script_end_procedure("Unable to send difference notice please review case.")
	
	EMReadScreen quarter, 1, 8, 14
	EMReadScreen IEVS_year, 4, 8, 22
	
	EMReadScreen Active_Programs, 13, 6, 68
	Active_Programs =trim(Active_Programs)

	EMReadScreen employer_info, 27, 8, 37
	employer_info = trim(employer_info)
	
	length = len(employer_info) 						'establishing the length of the variable
	position = InStr(employer_info, " AMT:")    		'sets the position at the deliminator                    
	employer_info = left(employer_info, length-position)	    'establishes employer as being before the delimiter

	EMReadScreen client_name, 35, 5, 24
    'Formatting the client name for the case note
    client_name = trim(client_name)                         'trimming the client name
    if instr(client_name, ",") then    						'Most cases have both last name and 1st name. This seperates the two names
        length = len(client_name)                           'establishing the length of the variable
        position = InStr(client_name, ",")                  'sets the position at the deliminator (in this case the comma)
        last_name = Left(client_name, position-1)           'establishes client last name as being before the deliminator
        first_name = Right(client_name, length-position)    'establishes client first name as after before the deliminator
    Else                                'In cases where the last name takes up the entire space, then the client name becomes the last name
        first_name = " "
        last_name = client_name
    END IF
    if instr(first_name, " ") then   						'If there is a middle initial in the first name, then it removes it
        length = len(first_name)                        	'trimming the 1st name
        position = InStr(first_name, " ")               	'establishing the length of the variable
        first_name = Left(first_name, position-1)       	'trims the middle initial off of the first name
    End if
	
	EMReadScreen Wage_option_type, 40, 7, 39
	Wage_option_type = trim(Wage_option_type) 'takes the period off of types when autofilled into dialog
	If right(Wage_option_type, 1) = "." THEN Wage_option_type = left(Wage_option_type, len(Wage_option_type) - 1) 
	
	EMReadScreen diff_date, 10, 14, 68
	diff_date = trim(diff_date)
	If diff_date = "" then				'sending the notice
        User_response = MsgBox ("Do you want to send Difference Notice?", vbquestion + vbyesno, "Send Difference Notice")
	    IF User_response = vbno then script_end_procedure ("Difference Notice has not been sent, IEVS match not acted on.")
	    IF User_response = vbyes then 
	    	EMwritescreen "015", 12, 46
	    	'sending the notice
	    	EMwritescreen "Y", 14, 37 'send Notice
	    	transmit		
	    	EMReadScreen edit_error, 2, 24, 2
	    	edit_error = trim(edit_error)
	    	if edit_error <> "" then script_end_procedure("Unable to send difference notice please review case.")  
	    END IF
		sent_notice = True
	Else 
		diff_sent = diff_sent + 1
		sent_notice = False
		PF3
		row = row + 1
	END IF 	
Loop until sent_notice = True

'----------------------------------------------------------------------------------------------------Logic for the case notes
'Updated IEVS_period to write into case note
If quarter = 1 then IEVS_quarter = "1st"
If quarter = 2 then IEVS_quarter = "2nd"
If quarter = 3 then IEVS_quarter = "3rd"
If quarter = 4 then IEVS_quarter = "4th"

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
 
Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days
IEVS_period = replace(IEVS_period, "/", " to ")

'----------------------------------------------------------------------------------------------------The case note
start_a_blank_CASE_NOTE
If IEVS_type = "WAGE" then Call write_variable_in_CASE_NOTE ("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH (" & first_name & ") DIFF NOTICE SENT-----")
If IEVS_type = "BEER" then Call write_variable_in_CASE_NOTE ("-----" & IEVS_year & " NON WAGE MATCH (B) (" & first_name & ") DIFF NOTICE SENT-----")
Call write_bullet_and_variable_in_CASE_NOTE("Period", IEVS_period)
Call write_bullet_and_variable_in_CASE_NOTE("Active Programs", programs)
Call write_bullet_and_variable_in_CASE_NOTE("Employer info:", employer_info)
Call write_variable_in_CASE_NOTE("----- ----- ----- ----- ----- ----- -----")
Call write_bullet_and_variable_in_CASE_NOTE("Type", Wage_option_type)
Call write_variable_in_CASE_NOTE("* Verification Requested: EVF and/or ATR")
Call write_bullet_and_variable_in_CASE_NOTE("Difference Notice Due", Due_date)
Call write_bullet_and_variable_in_CASE_NOTE("Date ATR on File", ATR_on_file_date)
Call write_variable_in_CASE_NOTE ("----- ----- ----- ----- ----- ----- -----")
Call write_variable_in_CASE_NOTE ("DEBT ESTABLISHMENT UNIT 612-348-4290 EXT 1-1-1")

'-------------------------------------------------------------------------------------------------------------CLAIM REFERRAL TRACKING FUNCTIONALITY 
IF Claim_Referral_checkbox = checked then

    BeginDialog Claim_Referral_Tracking, 0, 0, 301, 115, "CLAIM REFERRAL TRACKING FS & CASH ONLY"
      EditBox 60, 10, 75, 15, MAXIS_case_number
      DropListBox 195, 10, 95, 15, "Select One:"+chr(9)+"SNAP"+chr(9)+"MFIP"+chr(9)+"SNAP/MFIP", program_droplist
      EditBox 60, 30, 75, 15, Action_Date
      DropListBox 195, 30, 95, 15, "Select One:"+chr(9)+"Initial Claim Referral"+chr(9)+"Claim Determination", Action_Taken
      CheckBox 10, 55, 145, 10, "Sent Request for Additional Information ", Verif_Checkbox
      CheckBox 175, 55, 75, 10, "Overpayment Exists", Overpayment_Checkbox
      EditBox 60, 70, 230, 15, Other_Notes
      ButtonGroup ButtonPressed
        OkButton 185, 90, 50, 15
        CancelButton 240, 90, 50, 15
      Text 145, 35, 45, 10, "Action Taken:"
      Text 10, 75, 45, 10, "Other Notes:"
      Text 10, 35, 40, 10, "Action Date:"
      Text 145, 15, 40, 10, "Program(s):"
      Text 10, 15, 50, 10, "Case Number: "
    EndDialog
    Do
    	Do
    		err_msg = ""
    		dialog Claim_Referral_Tracking
    		IF buttonpressed = 0 then stopscript 
    		IF MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
    		IF isdate(action_date) = False then err_msg = err_msg & vbnewline & "* Enter a valid action date."
    		IF program_droplist = "Select One:" then err_msg = err_msg & vbnewline & "* Select a program."
    		IF Action_Taken = "Select One:" then err_msg = err_msg & vbnewline & "* Select an action."
    		IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "* Enter your worker signature."		
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
    	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in
    
    'Going to the MISC panel
    Call navigate_to_MAXIS_screen ("STAT", "MISC")
    Row = 6 
    EmReadScreen panel_number, 1, 02, 78 
    If panel_number = "0" then
    	EMWriteScreen "NN", 20,79
    	TRANSMIT	
    ELSE 
    	Do 
        	'Checking to see if the MISC panel is empty, if not it will find a new line'
        	EmReadScreen MISC_description, 25, row, 30
        	MISC_description = replace(MISC_description, "_", "")
        	If trim(MISC_description) = "" then 
    			PF9
        		EXIT DO
        	Else
                row = row + 1
        	End if
    	Loop Until row = 17
        If row = 17 then script_end_procedure("There is not a blank field in the MISC panel. Please delete a line(s), and run script again or update manually.")
    End if  
    EMReadScreen error_msg, 7, 24, 2
    'writing in the action taken and date to the MISC panel
    EMWriteScreen Action_Taken, Row, 30
    EMWriteScreen Action_Date, Row, 66 
    PF3
    'set TIKL------------------------------------------------------------------------------------------------------
    If Verif_Checkbox = checked then 
    	If Action_Taken = "Claim Determination" THEN 
    		Msgbox "You identified your case is ready to process for overpayment follow procedures for claim entry.  A TIKL will NOT be made."
    	ELSE 	
    		Call navigate_to_MAXIS_screen("DAIL", "WRIT")
    		call create_MAXIS_friendly_date(Action_Date, 10, 5, 18)
    		Call write_variable_in_TIKL("A potential overpayment exists on case. Please review case for reciept of requested information.")
    		PF3
    	END IF
    END IF	
    '-------------------------------------------------------------------------------------------------The case note
	'logic for the case note
	IF ATR_on_file_date = "" then ATR_on_file_date = "Not on file."
	
    start_a_blank_CASE_NOTE
    Call write_variable_in_case_note("-----CLAIM REFERRAL TRACKING - " & Action_Taken & "-----")
    Call write_bullet_and_variable_in_CASE_NOTE("Period", IEVS_period)
    Call write_bullet_and_variable_in_CASE_NOTE("Active Programs", program_droplist)
    Call write_bullet_and_variable_in_CASE_NOTE("Employer info", employer_info)
    Call write_variable_in_CASE_NOTE("----- ----- ----- ----- ----- ----- -----")
    Call write_bullet_and_variable_in_case_note("Action Date", Action_Date)
    Call write_bullet_and_variable_in_case_note("Other Notes", Other_Notes)
    Call write_variable_in_case_note("* Entries for these potential claims must be retained until further notice.")
    	IF (Verif_Checkbox = checked and Action_Taken <> "Claim Determination") then write_variable_in_case_note("* Additional verifications requested, TIKL set for 10 day return.")
    		IF Overpayment_Checkbox = checked then write_variable_in_case_note("* Overpayment exists, collection process to follow.")
    Call write_variable_in_CASE_NOTE ("----- ----- ----- ----- ----- ----- -----")
    Call write_variable_in_CASE_NOTE ("DEBT ESTABLISHMENT UNIT 612-348-4290 EXT 1-1-1")
    IF Overpayment_Checkbox = checked then 
    	script_end_procedure("You have indicated that an overpayment exists. Please follow the agency's procedure(s) for claim entry.")
    Else 
    	script_end_procedure ("Success!  A WAGE MATCH DIFFERENCE NOTICE HAS BEEN SENT." & vbnewline & vbnewline & "Please remember to send out an ATR/EVF from ECF.")
    END IF
Else 
	script_end_procedure("")
End if 