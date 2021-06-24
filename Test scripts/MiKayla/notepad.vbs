'GATHERING STATS===========================================================================================
name_of_script = "BULK - DEU-MATCH CLEARED.vbs"
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
		END IF
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine & "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine & "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
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
script_run_lowdown = ""
'TODO I need error proofing in multiple places on this script. in and out of IULA and IULB ensuring the case and on CCOL' 'need to check about adding for multiple claims'

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("06/17/2021", "GitHub #498 Updating the dialog box to ensure that a cleared method is entered.", "MiKayla Handley, Hennepin County")
call changelog_update("12/07/2019", "Added handling for coding the Excel spreadsheet. You must use BC, BE, BN, or CC only in the cleared status field.", "MiKayla Handley, Hennepin County")
call changelog_update("11/14/2017", "Program information will not be input into the Excel spreadsheet. This will not need to be added manually by staff completing the cases.", "Ilse Ferris, Hennepin County")
call changelog_update("06/05/2017", "Added handling for minor children in school (excluded income) & multiple people per case.", "Ilse Ferris, Hennepin County")
call changelog_update("03/20/2017", "Initial version.", "Ilse Ferris, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
discovery_date = date & ""
action_date = date & ""
'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to MAXIS
EMConnect ""

EMReadScreen mx_region, 10, 22, 48
If mx_region = "INQUIRY DB" Then
    continue_in_inquiry = MsgBox("It appears you are attempting to have the script clear these matches." & vbNewLine & vbNewLine & "However, you appear to be in MAXIS Inquiry." &vbNewLine & "*************************" & vbNewLine & "Do you want to continue?", vbQuestion + vbYesNo, "Confirm Inquiry")
    If continue_in_inquiry = vbNo Then script_end_procedure_with_error_report("Live script run was attempted in Inquiry and aborted.")
End If
'confirming that there is a worker signature on file.
If trim(worker_signature) = "" Then
    worker_signature = InputBox("How would you like to sign you case notes:", "Worker Signature")
End If
'dialog and dialog DO...Loop
DO
	DO
		'The dialog is defined in the loop as it can change as buttons are pressed
		'-------------------------------------------------------------------------------------------------DIALOG
		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 271, 180, "BULK-Match Cleared"
		  DropListBox 140, 15, 55, 15, "Select One:"+chr(9)+"BEER"+chr(9)+"BNDX"+chr(9)+"SDXS/SDXI"+chr(9)+"UNVI"+chr(9)+"UBEN"+chr(9)+"WAGE", select_match_type
		  DropListBox 140, 35, 120, 15, "Not Needed"+chr(9)+"Initial"+chr(9)+"Overpayment Exists"+chr(9)+"OP Non-Collectible (please specify)"+chr(9)+"No Savings/Overpayment", claim_referral_tracking_dropdown
		  EditBox 65, 55, 195, 15, other_notes
		  ButtonGroup ButtonPressed
		    PushButton 10, 115, 50, 15, "Browse:", select_a_file_button
		  EditBox 65, 115, 195, 15, IEVS_match_path
		  ButtonGroup ButtonPressed
		    OkButton 170, 160, 45, 15
		    CancelButton 220, 160, 45, 15
		  GroupBox 5, 5, 260, 70, "Complete prior to browsing the script:"
		  Text 10, 20, 120, 10, "Select the type of match to process:"
		  Text 10, 40, 130, 10, "Claim Referral Tracking on STAT/MISC:"
		  Text 10, 60, 45, 10, "Other Notes:"
		  GroupBox 5, 80, 260, 75, "Using the script:"
		  Text 10, 90, 250, 15, "Select the Excel file that contains the case information by selecting the 'Browse' button and locating the file."
		  Text 10, 135, 245, 15, "This script should be used when matches have been researched and ready to be cleared. "
		EndDialog
			err_msg = ""
			Dialog Dialog1
			cancel_confirmation
			If ButtonPressed = select_a_file_button then
				If IEVS_match_path <> "" then 'This is handling for if the BROWSE button is pushed more than once'
					objExcel.Quit 'Closing the Excel file that was opened on the first push'
					objExcel = "" 	'Blanks out the previous file path'
				END IF
				call file_selection_system_dialog(IEVS_match_path, ".xlsx") 'allows the user to select the file'
			END IF
			If select_match_type = "Select One:" then err_msg = err_msg & vbNewLine & "* Select type of match you are processing."
			If IEVS_match_path = "" then err_msg = err_msg & vbNewLine & "* Use the Browse Button to select the file that has your client data"
			If err_msg <> "" Then MsgBox err_msg
		LOOP UNTIL err_msg = ""
		If objExcel = "" Then call excel_open(IEVS_match_path, TRUE, TRUE, ObjExcel, objWorkbook)  'opens the selected excel file'
		If err_msg <> "" Then MsgBox err_msg
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = FALSE					'loops until user passwords back in

excel_row = 2			'establishing row to start
DO
	DO 'DAIL DO'
	    MAXIS_case_number	 = objExcel.Cells(excel_row, 1).Value    	 'CASE NUMBER
	    client_name 		 = objExcel.Cells(excel_row, 2).Value    	 'APPLICANT NAME
	    client_ssn 			 = objExcel.Cells(excel_row, 3).Value    	 'SSN
	    match_programs		 = objExcel.Cells(excel_row, 4).Value  		 'PROGRAM(S)
	    resolution_status	 = objExcel.Cells(excel_row, 7).Value    	 'RESOLUTION

	    'cleaned up
	    MAXIS_case_number 	= trim(MAXIS_case_number) 'remove extra spaces'
	    client_ssn 			= trim(client_ssn)
	    client_ssn 			= replace(client_ssn, "-", "") 'must be for IEVS to be used'
	    income_source 	   	= trim(income_source)
	    resolution_status  	= trim(resolution_status)
	    resolution_status  	= UCASE(resolution_status)
	    IF MAXIS_case_number = "" THEN EXIT DO 'goes to actions outside of do loop'
	    IF resolution_status = "" THEN
	    	EXIT DO 'goes to actions outside of do loop'
	    	match_found = FALSE
			match_cleared = "FALSE"
	    END IF
	    correct_resolution_status = FALSE
	    IF resolution_status = "BC" THEN correct_resolution_status = TRUE
	    IF resolution_status = "BE" THEN correct_resolution_status = TRUE
	    IF resolution_status = "BN" THEN correct_resolution_status = TRUE
	    IF resolution_status = "CC" THEN correct_resolution_status = TRUE

	    back_to_self
	    CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
 	    DO
 	    	EMReadScreen dail_check, 4, 2, 48
 	    	If next_dail_check <> "DAIL" then CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
 	    LOOP UNTIL dail_check = "DAIL"

  	    dail_row = 6			'Because the script brings each new case to the top of the page, dail_row starts at 6.
	    EMReadscreen number_of_dails, 1, 3, 67	'Reads where there count of dAILS is listed
 	    IF number_of_dails = " " Then
	    	exit do		'if this space is blank the rest of the DAIL reading is skipped
	    	match_cleared = "no DAIL"
	    END IF
	    EMReadScreen DAIL_case_number, 8, dail_row, 73
 	    DAIL_case_number = trim(DAIL_case_number)
 	    If DAIL_case_number <> MAXIS_case_number then exit do
 	       'Determining if there is a new case number...
	    msgbox maxis_case_number & vbcr & DAIL_case_number
	    EMReadScreen new_case, 8, dail_row, 63
 	       new_case = trim(new_case)
 	       IF new_case <> "CASE NBR" THEN '...if there is NOT a new case number, the script will read the DAIL type, month, year, and message...
 	        	Call write_value_and_transmit("T", dail_row, 3)
 	        	dail_row = 6
 	       ELSEIF new_case = "CASE NBR" THEN
 	        '...if the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
 	        	Call write_value_and_transmit("T", dail_row + 1, 3)
 	        	dail_row = 6
 	       END IF

 	    EMReadScreen DAIL_type, 4, dail_row, 6
	    DAIL_type = trim(DAIL_type)
	    IF select_match_type = DAIL_type THEN match_found = TRUE

	    EMReadScreen dail_msg, 61, dail_row, 20 'future proofing'
 	       dail_msg = trim(dail_msg)

 	       EMReadScreen dail_month, 8, dail_row, 11
 	       dail_month = trim(dail_month)

 	       dail_row = dail_row + 1
	    EMReadScreen dail_msg, 4, row, 6
	    msgbox dail_msg & vbcr & select_match_type
	    If trim(dail_msg) <> trim(select_match_type) then
	    	match_found = False
	    	row = row + 1
	    	EMReadScreen new_case, 9, row, 63
	        If new_case = "CASE NBR:" then
	        	EMreadscreen case_number, 7, row, 73
	        	If trim(case_number) = MAXIS_case_number then
	        		row = row + 1
	        	Else
	        		exit do
	        	End if
	        Else
	        	msgbox "1." & MAXIS_case_number & vbcr & "new_case" & new_case & vbcr & "row: " & row & vbcr & "match found: " & match_found
	        End if
            If row = 19 then
            	PF8
            	row = 6
            End if
        Else
            EMReadScreen client_social, 9, row, 20
	    	If client_social <> Client_SSN then
	    		match_found = False
	    		row = row + 1
	    		msgbox "2." & MAXIS_case_number & vbcr & "row: " & row & vbcr & "match found: " & match_found
	    	Else
	    		match_found = true
	    		msgbox "3." & MAXIS_case_number & vbcr & "row: " & row & vbcr & "match found: " & match_found
	    		exit do
	    	End if
	    End if
	LOOP

	If match_found = False then
		match_found = False 'no case note'
		objExcel.cells(excel_row, 10).value = "A IEVS match wasn't found on DAIL/DAIL or SSN did not match."
	End if


'----------------------------------------------------------------------------------------------------IEVS
If match_found = True then
'Navigating deeper into the match interface
CALL write_value_and_transmit("I", row, 3)   'navigates to INFC
CALL write_value_and_transmit("IEVP", 20, 71)   'navigates to IEVP
EMReadScreen error_msg, 7, 24, 2
If error_msg = "NO IEVS" then 'checking for error msg'
	objExcel.cells(excel_row, 10).value = "No IEVS matches found for SSN " & Client_SSN & "/Could not access IEVP."
	match_found = False
END IF
match_found = TRUE
row = 7
'Ensuring that match has not already been resolved.
Do
	EMReadScreen days_pending, 5, row, 72
	days_pending = trim(days_pending)
	If IsNumeric(days_pending) = false then
		objExcel.cells(excel_row, 10).value = "No pending IEVS match found. Please review IEVP."
		match_found = False
		exit do
	END IF
LOOP
'---------------------------------------------------------------------Reading potential errors for out-of-county cases
    CALL write_value_and_transmit("U", row, 3)   'navigates to IULA
    EMReadScreen OutOfCounty_error, 12, 24, 2
	IF OutOfCounty_error = "MATCH IS NOT" THEN
    	script_end_procedure_with_error_report("Out-of-county case. Cannot update.")
    ELSE
    	EMReadScreen number_IEVS_type, 3, 7, 12 'read the match type'
        IF number_IEVS_type = "A30" THEN match_type = "BNDX"
        IF number_IEVS_type = "A40" THEN match_type = "SDXS/I"
        IF number_IEVS_type = "A70" THEN match_type = "BEER"
        IF number_IEVS_type = "A80" THEN match_type = "UNVI"
        IF number_IEVS_type = "A60" THEN match_type = "UBEN"
        IF number_IEVS_type = "A50" THEN match_type = "WAGE"
		IF number_IEVS_type = "A51" THEN match_type = "WAGE"
        IEVS_year = ""
		IF match_type = "WAGE" THEN
        	EMReadScreen select_quarter, 1, 8, 14
        	EMReadScreen IEVS_year, 4, 8, 22
        ELSEIF match_type = "UBEN" THEN
        	EMReadScreen IEVS_month, 2, 5, 68
        	EMReadScreen IEVS_year, 4, 8, 71
        ELSEIF match_type = "BEER" THEN
        	EMReadScreen IEVS_year, 2, 8, 15
        	IEVS_year = "20" & IEVS_year
    	ELSEIF match_type = "UNVI" THEN
    		EMReadScreen IEVS_year, 4, 8, 15
    		msgbox IEVS_year
    		select_quarter = "YEAR"
    	END IF
	END IF
    EMReadScreen number_IEVS_type, 3, 7, 12 'read the DAIL msg'
    IF number_IEVS_type = "A30" THEN match_type = "BNDX"
    IF number_IEVS_type = "A40" THEN match_type = "SDXS/I"
    IF number_IEVS_type = "A70" THEN match_type = "BEER"
    IF number_IEVS_type = "A80" THEN match_type = "UNVI"
    IF number_IEVS_type = "A60" THEN match_type = "UBEN"
    IF number_IEVS_type = "A50" THEN match_type = "WAGE"
	IF number_IEVS_type = "A51" THEN match_type = "WAGE"
    '--------------------------------------------------------------------Client name
    EMReadScreen panel_name, 4, 02, 52
    IF panel_name <> "IULA" THEN
		match_found = FALSE
		EXIT DO
	END IF
    EMReadScreen client_name, 35, 5, 24
    client_name = trim(client_name)                         'trimming the client name
    IF instr(client_name, ",") THEN    						'Most cases have both last name and 1st name. This separates the two names
    	length = len(client_name)                           'establishing the length of the variable
    	position = InStr(client_name, ",")                  'sets the position at the deliminator (in this case the comma)
    	last_name = Left(client_name, position-1)           'establishes client last name as being before the deliminator
    	first_name = Right(client_name, length-position)    'establishes client first name as after before the deliminator
    ELSEIF instr(first_name, " ") THEN   					'If there is a middle initial in the first name, THEN it removes it
    	length = len(first_name)                        	'trimming the 1st name
    	position = InStr(first_name, " ")               	'establishing the length of the variable
    	first_name = Left(first_name, position-1)       	'trims the middle initial off of the first name
    ELSE                                					'In cases where the last name takes up the entire space, THEN the client name becomes the last name
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
    IF len(active_programs) <> 1 THEN
    	match_found = FALSE
		match_cleared = "Not cleared - programs"
    	EXIT DO
    END IF
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
	'----------------------------------------------------------------------------------------------------notice sent
	EMReadScreen notice_sent, 1, 14, 37
	EMReadScreen sent_date, 8, 14, 68
	sent_date = trim(sent_date)
	IF sent_date = "" THEN sent_date = "N/A"
	IF sent_date <> "" THEN sent_date = replace(sent_date, " ", "/")
	EMReadScreen clear_code, 2, 12, 58
    '----------------------------------------------------------------------------------------------------Employer info & difference notice info
    IF match_type = "UBEN" THEN income_source = "Unemployment"
    IF match_type = "UNVI" THEN income_source = "NON-WAGE"
    IF match_type = "WAGE" THEN
		EMReadScreen income_source, 50, 8, 37 'was 37' should be to the right of employer and the left of amount
		income_source = trim(income_source)
		length = len(income_source)		'establishing the length of the variable
		'should be to the right of employer and the left of amount '
	    IF instr(income_source, " AMOUNT: $") THEN
	    	position = InStr(income_source, " AMOUNT: $")    		      'sets the position at the deliminator
	    	income_source = Left(income_source, position)  'establishes employer as being before the deliminator
	    Elseif instr(income_source, " AMT: $") THEN 					  'establishing the length of the variable
	    	position = InStr(income_source, " AMT: $")    		      'sets the position at the deliminator
	    	income_source = Left(income_source, position)  'establishes employer as being before the deliminator
	    END IF
    END IF
    IF match_type = "BEER" THEN
    	EMReadScreen income_source, 50, 8, 28 'was 37' should be to the right of employer and the left of amount
    	income_source = trim(income_source)
    	length = len(income_source)		'establishing the length of the variable
    	'should be to the right of employer and the left of amount '
    	IF instr(income_source, " AMOUNT: $") THEN
    		position = InStr(income_source, " AMOUNT: $")    		      'sets the position at the deliminator
    		income_source = Left(income_source, position)  'establishes employer as being before the deliminator
    	Elseif instr(income_source, " AMT: $") THEN 					  'establishing the length of the variable
    		position = InStr(income_source, " AMT: $")    		      'sets the position at the deliminator
    		income_source = Left(income_source, position)  'establishes employer as being before the deliminator
    	END IF
    END IF
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
    	EMReadScreen cleared_field, 2, row + 1, col + 1
    	If cleared_field <> "__" then
        	match_found = "Unable to update cleared status on IULA."
        Else
        	EMWriteScreen resolution_status, row + 1, col + 1
        END IF
    Next
        msgbox "did we clear anything?"
        CALL write_value_and_transmit("10", 12, 46)   'navigates to IULB
        'resolved notes depending on the resolution_status
        IF resolution_status = "BC" then CALL write_value_and_transmit("Case closed " & other_notes , 8, 6)   'BC
        IF resolution_status = "BE" then CALL write_value_and_transmit("No change " & other_notes , 8, 6)   'BE
        If resolution_status = "BN" then CALL write_value_and_transmit("Already known - No savings " & other_notes , 8, 6)   'BN
        If resolution_status = "CC" then CALL write_value_and_transmit("Claim entered", 8, 6)   'CC
        match_found = "match cleared"
        IF match_type = "BEER" THEN match_type_letter = "B"
        IF match_type = "UBEN" THEN match_type_letter = "U"
        IF match_type = "UNVI" THEN match_type_letter = "U"
        IF match_type = "WAGE" THEN
        IF select_quarter = 1 THEN IEVS_quarter = "1ST"
        IF select_quarter = 2 THEN IEVS_quarter = "2ND"
        IF select_quarter = 3 THEN IEVS_quarter = "3RD"
        IF select_quarter = 4 THEN IEVS_quarter = "4TH"
    	IEVS_period = trim(IEVS_period)
    	IF match_type <> "UBEN" THEN IEVS_period = replace(IEVS_period, "/", " to ")
    	IF match_type = "UBEN" THEN IEVS_period = replace(IEVS_period, "-", "/")

  	    IF claim_referral_tracking_dropdown <> "Not Needed" THEN
  	        start_a_blank_case_note
  	        IF claim_referral_tracking_dropdown =  "Initial" THEN
		    	CALL write_variable_in_case_note("Claim Referral Tracking - Initial")
		    ELSE
		    	CALL write_variable_in_case_note("Claim Referral Tracking - " & MISC_action_taken)
		    END IF
		    CALL write_bullet_and_variable_in_case_note("Action Date", action_date)
		    CALL write_bullet_and_variable_in_case_note("Active Program(s)", programs)
		    CALL write_variable_in_case_note("* Entries for these potential claims must be retained until further notice.")
		    IF case_note_only = TRUE THEN CALL write_variable_in_case_note("Maxis case is inactive unable to add or update MISC panel")
		    CALL write_variable_in_case_note("-----")
		    CALL write_variable_in_case_note(worker_signature)
		    PF3
		END IF

		start_a_blank_case_note
		IF match_type = "WAGE" THEN CALL write_variable_in_case_note("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH"  & " (" & first_name & ") " & cleared_header & header_note & "-----")
		IF match_type = "BEER" THEN CALL write_variable_in_case_note("-----" & IEVS_year & " NON-WAGE MATCH(" & match_type_letter & ")" & " (" & first_name & ") " & cleared_header & header_note & "-----")
		IF match_type = "UNVI" THEN CALL write_variable_in_case_note("-----" & IEVS_year & " NON-WAGE MATCH(" & match_type_letter & ")" & " (" & first_name & ") " & cleared_header & header_note & "-----")
		IF match_type = "UBEN" THEN CALL write_variable_in_case_note("-----" & IEVS_period & " NON-WAGE MATCH(" & match_type_letter & ")" & " (" & first_name & ") " & cleared_header & header_note & "-----")
		IF match_type = "BNDX" THEN CALL write_variable_in_case_note("-----" & IEVS_period & " NON-WAGE MATCH(" & match_type & ")" & " (" & first_name & ") " & cleared_header & header_note & "-----")
		CALL write_bullet_and_variable_in_case_note("Discovery date", discovery_date)
		CALL write_bullet_and_variable_in_case_note("Period", IEVS_period)
		CALL write_bullet_and_variable_in_case_note("Active Programs", programs)
		CALL write_bullet_and_variable_in_case_note("Source of income", income_source)
		CALL write_variable_in_case_note("----- ----- ----- ----- ----- ----- -----")
		CALL write_bullet_and_variable_in_case_note("Date Diff notice sent", sent_date)
		IF resolution_status = "BC" THEN CALL write_variable_in_case_note("* Case closed.")
		IF resolution_status = "BE" THEN CALL write_variable_in_case_note("* No Overpayments or savings were found related to this match.")
		IF resolution_status = "BN" THEN CALL write_variable_in_case_note("* Client reported income. Correct income is in JOBS/BUSI and budgeted.")
		CALL write_bullet_and_variable_in_case_note("Other Notes", other_notes)
		CALL write_variable_in_case_note("----- ----- ----- ----- ----- ----- -----")
		CALL write_variable_in_case_note("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1")
		PF3 'to save casenote'

        objExcel.cells(excel_row, 8).value = date_cleared
	    objExcel.cells(excel_row, 9).value = match_found
	END IF
	excel_row = excel_row + 1
	STATS_counter = STATS_counter + 1
LOOP UNTIL objExcel.Cells(excel_row, 1).value = ""

match_found = ""
'Excel headers and formatting the columns
objExcel.Cells(1, 1).Value     = "CASE NUMBER"
objExcel.Cells(1, 2).Value     = "APPLICANT NAME"
objExcel.Cells(1, 3).Value     = "SSN"
objExcel.Cells(1, 4).Value     = "PROGRAM(S)"
objExcel.Cells(1, 5).Value     = "AMOUNT"
objExcel.Cells(1, 6).Value     = "SOURCE OF INCOME"
objExcel.Cells(1, 7).Value     = "RESOLUTION"
objExcel.Cells(1, 8).Value     = "DATE CLEARED"
objExcel.Cells(1, 9).Value     = "NOTES"

FOR i = 1 to 10		'formatting the cells
    objExcel.Cells(1, i).Font.Bold 			= TRUE		'bold font'
    objExcel.Columns(i).AutoFit()						'sizing the columns'
	objExcel.Columns(i).HorizontalAlignment = -4131		'Centers the text for the columns with days remaining and difference notice
	objExcel.Columns(8).NumberFormat = "mm/dd/yy"		'formats the date column as MM/DD/YY
NEXT

STATS_counter = STATS_counter - 1						'removes 1 to correct the count
script_end_procedure_with_error_report("Success! The IEVS match cases have now been updated. Please review the NOTES section to review the cases/follow up work to be completed.")
END IF
