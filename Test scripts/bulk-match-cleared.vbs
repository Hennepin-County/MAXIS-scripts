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
call changelog_update("12/07/2019", "Added handling for coding the Excel spreadsheet. You must use BC, BE, BN, or CC only in the cleared status field.", "MiKayla Handley, Hennepin County")
call changelog_update("11/14/2017", "Program information will not be input into the Excel spreadsheet. This will not need to be added manually by staff completing the cases.", "Ilse Ferris, Hennepin County")
call changelog_update("06/05/2017", "Added handling for minor children in school (excluded income) & multiple people per case.", "Ilse Ferris, Hennepin County")
call changelog_update("03/20/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display

'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to MAXIS
EMConnect ""

'dialog and dialog DO...Loop
Do
	Do
		'The dialog is defined in the loop as it can change as buttons are pressed
		'-------------------------------------------------------------------------------------------------DIALOG
		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 266, 130, "BULK-Match Cleared"
		  ButtonGroup ButtonPressed
		    PushButton 15, 45, 50, 15, "Browse:", select_a_file_button
		  EditBox 70, 45, 180, 15, IEVS_match_path
		  EditBox 60, 90, 195, 15, Edit2
		  CheckBox 10, 110, 115, 15, "Select for claim referral tracking", claim_referral_tracking_checkbox
		  ButtonGroup ButtonPressed
		    OkButton 165, 110, 45, 15
		    CancelButton 215, 110, 45, 15
		  GroupBox 5, 5, 250, 80, "Using the script:"
		  Text 15, 65, 230, 15, "Select the Excel file that contains the case information by selecting the 'Browse' button and locate the file."
		  Text 10, 95, 45, 10, "Other Notes:"
		  Text 15, 20, 235, 20, "This script should be used when matches have been researched and ready to be cleared. "
		EndDialog
			err_msg = ""
			Dialog Dialog1
			cancel_confirmation
			If ButtonPressed = select_a_file_button then
				If IEVS_match_path <> "" then 'This is handling for if the BROWSE button is pushed more than once'
					objExcel.Quit 'Closing the Excel file that was opened on the first push'
					objExcel = "" 	'Blanks out the previous file path'
				End If
				call file_selection_system_dialog(IEVS_match_path, ".xlsx") 'allows the user to select the file'
			End If
			If match_type = "Select One:" then err_msg = err_msg & vbNewLine & "* Select type of match you are processing."
			If IEVS_match_path = "" then err_msg = err_msg & vbNewLine & "* Use the Browse Button to select the file that has your client data"
			If err_msg <> "" Then MsgBox err_msg
		Loop until err_msg = ""
		If objExcel = "" Then call excel_open(IEVS_match_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
		If err_msg <> "" Then MsgBox err_msg
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in


    'Excel headers and formatting the columns
    objExcel.Cells(1, 1).Value     = "CASE NUMBER"
    objExcel.Cells(1, 1).Font.Bold = TRUE
    objExcel.Cells(1, 2).Value     = "APPLICANT NAME"
    objExcel.Cells(1, 2).Font.Bold = TRUE
    objExcel.Cells(1, 3).Value     = "SSN"
    objExcel.Cells(1, 3).Font.Bold = TRUE
    objExcel.Cells(1, 4).Value     = "AMOUNT"
    objExcel.Cells(1, 4).Font.Bold = TRUE
    objExcel.Cells(1, 5).Value     = "INCOME"
    objExcel.Cells(1, 5).Font.Bold = TRUE
    objExcel.Cells(1, 6).Value     = "RESOLUTION"
    objExcel.Cells(1, 6).Font.Bold = TRUE
    objExcel.Cells(1, 7).Value     = "NOTICE DATE"
    objExcel.Cells(1, 7).Font.Bold = TRUE
    objExcel.Cells(1, 8).Value     = "NOTES"
    objExcel.Cells(1, 8).Font.Bold = TRUE

	excel_row = 2			'establishing row to start
DO
	MAXIS_case_number 	= objExcel.cells(excel_row, 1).value	'establishes MAXIS case number
	client_SSN 			= objExcel.cells(excel_row, 3).value	'establishes client SSN
	'income_source		= objExcel.cells(excel_row, 6).value	'establishes employer name
	cleared_status	    = objExcel.cells(excel_row, 8).value	'establishes cleared status for the match
	'cleaned up
	MAXIS_case_number 	= trim(MAXIS_case_number) 'remove extra spaces'
	client_SSN 			= trim(client_SSN)
	client_SSN 			= replace(client_SSN, "-", "")
	income_source 	   	= trim(income_source)
	cleared_status 	  	= trim(cleared_status)

	IF cleared_status <> "BC" or cleared_status <> "BE" or cleared_status <> "BN" or cleared_status <> "CC" THEN err_msg = err_msg & vbNewLine & "Please only use BC, BE, BN, or CC when clearing a match."
    If MAXIS_case_number = "" THEN exit do 'goes to actions outside of do loop'
	back_to_self
	'----------------------------------------------------------------------------------------------------DAIL
	Call navigate_to_MAXIS_screen("DAIL", "DAIL")
	'Making sure that the user is on an acceptable DAIL message
	EMReadScreen case_number, 8, 5, 73
	case_number = trim(case_number)
	IF case_number <> MAXIS_case_number then
		EMreadscreen case_number, 8, 7, 73   'DAILS often read down two check to see if matching'
		 If case_number <> MAXIS_case_number then
			objExcel.cells(excel_row, 9).value = "Case number errror."
			match_found = False
			case_note_actions = FALSE
		End if
	Else
	row = 6    'establishing 1st row to search
	Do
		EMReadScreen IEVS_message, 4, row, 6
		'msgbox IEVS_message & vbcr & match_type
		If IEVS_message <> match_type then
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
				'msgbox "1." & MAXIS_case_number & vbcr & "new_case" & new_case & vbcr & "row: " & row & vbcr & "match found: " & match_found
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
				'msgbox "2." & MAXIS_case_number & vbcr & "row: " & row & vbcr & "match found: " & match_found
			Else
				match_found = true
				'msgbox "3." & MAXIS_case_number & vbcr & "row: " & row & vbcr & "match found: " & match_found
				exit do
			End if
		End if
	Loop until match_found = true
	If match_found = False then
		case_note_actions = False 'no case note'
		objExcel.cells(excel_row, 9).value = "A IEVS match wasn't found on DAIL/DAIL or SSN did not match."
	End if
	End if

	'----------------------------------------------------------------------------------------------------IEVS
	If match_found = True then
	    'Navigating deeper into the match interface
	    CALL write_value_and_transmit("I", row, 3)   'navigates to INFC
	    CALL write_value_and_transmit("IEVP", 20, 71)   'navigates to IEVP
		EMReadScreen error_msg, 7, 24, 2
		If error_msg = "NO IEVS" then 'checking for error msg'
			objExcel.cells(excel_row, 9).value = "No matches found for SSN " & client_SSN & "/Could not access IEVP."
			case_note_actions = False
		Else
			row = 7
		    'Ensuring that match has not already been resolved.
		    Do
				EMReadScreen number_IEVS_type, 3, row, 41 'read the match type'
				IF number_IEVS_type = "A30" THEN match_type = "BNDX"
				IF number_IEVS_type = "A40" THEN match_type = "SDXS/I"
				IF number_IEVS_type = "A70" THEN match_type = "BEER"
				IF number_IEVS_type = "A80" THEN match_type = "UNVI"
				IF number_IEVS_type = "A60" THEN match_type = "UBEN"
				IF number_IEVS_type = "A50" or number_IEVS_type = "A51"  THEN match_type = "WAGE"

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

				EMReadScreen days_pending, 5, row, 72
		    	days_pending = trim(days_pending)
		    	If IsNumeric(days_pending) = false then
					objExcel.cells(excel_row, 9).value = "No pending IEVS match found. Please review IEVP."
					case_note_actions = False
					exit do
		    	ELSE
	           		row = row + 1
				END IF
			Loop until row = 17

			If case_note_actions <> True then
				If match_type = "WAGE" then
			    	objExcel.cells(excel_row, 9).value = "This WAGE match is not for a quarter. Please process manually."
			    Elseif match_type = "BEER" then
					objExcel.cells(excel_row, 9).value = "This BEER match is not for a year. Please process manually."
				END if
				case_note_actions = False
			Else
			'---------------------------------------------------------------------Reading potential errors for out-of-county cases
				CALL write_value_and_transmit("U", row, 3)   'navigates to IULA
				EMReadScreen OutOfCounty_error, 12, 24, 2
				IF OutOfCounty_error = "MATCH IS NOT" then
				script_end_procedure_with_error_report("Out-of-county case. Cannot update.")

				'--------------------------------------------------------------------Client name
			    EmReadScreen panel_name, 4, 02, 52
			    IF panel_name <> "IULA" THEN script_end_procedure_with_error_report("Script did not find IULA.")
			    EMReadScreen client_name, 35, 5, 24
			    client_name = trim(client_name)                         'trimming the client name
			    IF instr(client_name, ",") THEN    						'Most cases have both last name and 1st name. This separates the two names
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
			    IF match_type = "UBEN" THEN income_source = "Unemployment"
			    IF match_type = "UNVI" THEN income_source = "NON-WAGE"
			    IF match_type = "WAGE" THEN
				    EMReadScreen income_source, 50, 8, 37 'was 37' should be to the right of emplyer and the left of amount
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
					EMReadScreen income_source, 50, 8, 28 'was 37' should be to the right of emplyer and the left of amount
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
						objExcel.cells(excel_row, 9).value = "Unable to update cleared status on IULA."
						case_note_actions = False
					Else
						EMWriteScreen cleared_status, row + 1, col + 1
					End if
				Next

                CALL write_value_and_transmit("10", 12, 46)   'navigates to IULB

			    'resolved notes depending on the cleared_status
			    If cleared_status = "BC" then CALL write_value_and_transmit("Case closed.", 8, 6)   'BC
                If cleared_status = "BE" then CALL write_value_and_transmit("No change.", 8, 6)   'BE
			    If cleared_status = "BN" then CALL write_value_and_transmit("Already known - No savings.", 8, 6)   'BN
			    If cleared_status = "CC" then CALL write_value_and_transmit("Claim entered.", 8, 6)   'CC
				objExcel.cells(excel_row, 9).value = "IEVS match cleared"
                case_note_actions = True
				End if
			End if
		End if
	End if

    If case_note_actions = True then		'Formatting for the case note

	    IF match_type = "BEER" THEN match_type_letter = "B"
	    IF match_type = "UBEN" THEN match_type_letter = "U"
	    IF match_type = "UNVI" THEN match_type_letter = "U"

	    verifcation_needed = trim(verifcation_needed) 	'takes the last comma off of verifcation_needed when autofilled into dialog if more more than one app date is found and additional app is selected
	    IF right(verifcation_needed, 1) = "," THEN verifcation_needed = left(verifcation_needed, len(verifcation_needed) - 1)
	    IF match_type = "WAGE" THEN
	    	IF select_quarter = 1 THEN IEVS_quarter = "1ST"
	    	IF select_quarter = 2 THEN IEVS_quarter = "2ND"
	    	IF select_quarter = 3 THEN IEVS_quarter = "3RD"
	    	IF select_quarter = 4 THEN IEVS_quarter = "4TH"
	    END IF

	    IEVS_period = trim(IEVS_period)
	    IF match_type <> "UBEN" THEN IEVS_period = replace(IEVS_period, "/", " to ")
	    IF match_type = "UBEN" THEN IEVS_period = replace(IEVS_period, "-", "/")
	    Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days

	    'adding specific wording for case note header for each cleared status
	    If cleared_status = "BC" then cleared_header_info = " (" & first_name & ") CLEARED BC-CASE CLOSED"
	    If cleared_status = "BE" then cleared_header_info = " (" & first_name & ") CLEARED BE-NO CHANGE"
	    If cleared_status = "BN" then cleared_header_info = " (" & first_name & ") CLEARED BN-KNOWN"
	    If cleared_status = "CC" then cleared_header_info = " (" & first_name & ") CLEARED CC-CLAIM ENTERED"

		'Case noting the actions taken
        start_a_blank_CASE_NOTE
        If match_type = "WAGE" then Call write_variable_in_CASE_NOTE("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH" & cleared_header_info & "-----")
		If match_type = "BEER" then Call write_variable_in_CASE_NOTE("-----" & IEVS_year & "  NON-WAGE MATCH (B)" & cleared_header_info & "-----")
		Call write_bullet_and_variable_in_CASE_NOTE("Period", IEVS_period)
		Call write_bullet_and_variable_in_CASE_NOTE("Active Programs", programs)
        Call write_bullet_and_variable_in_CASE_NOTE("Source of income", income_source)
        call write_variable_in_CASE_NOTE("------ ----- -----")
        If cleared_status = "BN" or cleared_status = "BE" then
			Call write_variable_in_CASE_NOTE("* Client reported income. Correct income is in JOBS/BUSI and budgeted")
			Call write_variable_in_CASE_NOTE("* This income was known and budgeted prior to COVID 19")
		END IF
        If cleared_status <> "CC" then Call write_variable_in_CASE_NOTE("* No collectible overpayments or savings were found related to this match.")
        call write_variable_in_CASE_NOTE("------ ----- ----- ----- -----")
        Call write_variable_in_CASE_NOTE("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1")
    End if

	excel_row = excel_row + 1
	MAXIS_case_number = ""
	client_SSN = ""
	STATS_counter = STATS_counter + 1

LOOP UNTIL objExcel.Cells(excel_row, 1).value = ""	'looping until the list of cases to check for recert is complete\
'Centers the text for the columns with days remaining and difference notice

objExcel.Columns(1).HorizontalAlignment = -4131
objExcel.Columns(2).HorizontalAlignment = -4131
objExcel.Columns(3).HorizontalAlignment = -4131
objExcel.Columns(4).HorizontalAlignment = -4131
objExcel.Columns(5).HorizontalAlignment = -4131
objExcel.Columns(6).HorizontalAlignment = -4131
objExcel.Columns(7).HorizontalAlignment = -4131
objExcel.Columns(8).HorizontalAlignment = -4131

'Formatting the column width.
FOR i = 1 to 9
	objExcel.Columns(i).AutoFit()
NEXT
'add pf3 at the end of the run and error handling for blank cleared status'
STATS_counter = STATS_counter - 1		'removes 1 to correct the count
script_end_procedure_with_error_report("Success! The IEVS match cases have now been updated. Please review the NOTES section to review the cases/follow up work to be completed.")
