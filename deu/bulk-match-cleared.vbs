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

'THE SCRIPT-----------------------------------------------------------------------------------------------------------
EMConnect ""
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr
back_to_self 'resetting MAXIS back to self before getting started
Call MAXIS_footer_month_confirmation	'ensuring we are in the correct footer month/year

Do
	Do
		'The dialog is defined in the loop as it can change as buttons are pressed
	    '-------------------------------------------------------------------------------------------------DIALOG
	    Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 271, 240, "BULK-Match Cleared"
		  DropListBox 140, 15, 55, 15, "Select One:"+chr(9)+"BEER"+chr(9)+"BNDX"+chr(9)+"SDXS/SDXI"+chr(9)+"UNVI"+chr(9)+"UBEN"+chr(9)+"WAGE", match_type
		  DropListBox 140, 35, 120, 15, "Not Needed"+chr(9)+"Initial"+chr(9)+"Overpayment Exists"+chr(9)+"OP Non-Collectible (please specify)"+chr(9)+"No Savings/Overpayment", claim_referral_tracking_dropdown
		  EditBox 65, 55, 195, 15, other_notes
		  ButtonGroup ButtonPressed
		    PushButton 10, 115, 50, 15, "Browse:", select_a_file_button
		  EditBox 65, 115, 195, 15, IEVS_match_path
		  ButtonGroup ButtonPressed
		    PushButton 115, 175, 145, 15, "Open IEVS Template Excel File", open_ievs_template_file_button
		    OkButton 170, 220, 45, 15
		    CancelButton 220, 220, 45, 15
		  GroupBox 5, 5, 260, 70, "Complete prior to browsing the script:"
		  Text 10, 20, 120, 10, "Select the type of match to process:"
		  Text 10, 40, 130, 10, "Claim Referral Tracking on STAT/MISC:"
		  Text 10, 60, 45, 10, "Other Notes:"
		  GroupBox 5, 80, 260, 135, "Using the script:"
		  Text 10, 90, 250, 15, "Select the Excel file that contains the case information by selecting the 'Browse' button and locating the file."
		  Text 10, 135, 245, 15, "This script should be used when matches have been researched and ready to be cleared. "
		  Text 10, 155, 245, 20, "You MUST use the correct Excel layout for this script to work properly. The column positions and layout can be found in the IEVS Template Excel file."
		  Text 10, 195, 245, 20, "If you use a different layout in the file you select, the script will likely not function correctly."
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
		If match_type = "Select One:" then err_msg = err_msg & vbNewLine & "* Select type of match you are processing."
		If IEVS_match_path = "" then err_msg = err_msg & vbNewLine & "* Use the Browse Button to select the file that has your client data"
		If ButtonPressed = open_ievs_template_file_button Then
			err_msg = "LOOP"
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/BlueZone%20Script%20Resources/IEVS%20TEMPLATE.xlsx"
		End If
		If err_msg <> "" and err_msg <> "LOOP" Then MsgBox err_msg
	LOOP UNTIL err_msg = ""
		If objExcel = "" Then call excel_open(IEVS_match_path, TRUE, TRUE, ObjExcel, objWorkbook)  'opens the selected excel file'
		If err_msg <> "" Then MsgBox err_msg
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = FALSE					'loops until user passwords back in
'setting the footer month to make the updates in'

'Establishing array
DIM match_based_array()  'Declaring the array this is what this list is
ReDim match_based_array(other_note_const, 0)  'Resizing the array 'that ,list is going to have 20 parameter but to start with there is only one paparmeter it gets complicated - grid'
'for each row the column is going to be the same information type
'Creating constants to value the array elements this is why we create constants
const date_posted_to_maxis_const	 	= 0 '=  Date Posted to Maxis'
const worker_number_const			 	= 1 '=  Basket
const client_DOB_const 				 	= 2 '=  DOB
const relationship_const			 	= 3 '=  Relationship
const case_earner_name_const	     	= 4 '=  Earner Name
const maxis_case_number_const   		= 5 '=  Case #
const client_name_const					= 6 '=  Name
const client_ssn_const					= 7 '=  SSN
const program_const  				 	= 8 '=  Prog
const amount_const 					 	= 9 '=  Amount
const income_source_const		     	= 10 '=  Employer
const notice_sent_const		     		= 11 '=  Date Notice Sent
const notice_sent_date_const		    = 12 '=  Date Notice Sent
const resolution_status_const   	 	= 13 '=  How cleared
const IEVS_period_const					= 14
const date_cleared_const			 	= 15 '=  Date cleared
const assigned_to_const				 	= 16 '=  Who worker who cleared
const match_cleared_const				= 17 '=  true/false
const other_note_const					= 18 '=  case note to check match cleared

'setting the columns - using constant so that we know what is going on'
const excel_col_date_posted_to_maxis	 = 1 'A' 'Date Posted to Maxis'
const excel_col_worker_number 			 = 2 'B'  Worker #
const excel_col_client_DOB				 = 3 'C' DOB
const excel_col_relationship			 = 4 'D' Relationship
const excel_col_case_earner_name 	     = 5 'E' Earner Name
const excel_col_case_number   			 = 6 'F' Case Number
const excel_col_client_name				 = 7 'G' Name
const excel_col_client_ssn				 = 8 'H' SSN
const excel_col_program  				 = 9 'I' Prog
const excel_col_amount 					 = 10 'J' Amount
const excel_col_income_source		     = 11 'K' Employer
const excel_date_notice_sent		     = 12 'L' Date Notice Sent
const excel_col_resolution_status   	 = 13 'M' How cleared
const excel_col_date_cleared			 = 14 'N' Date cleared
const excel_col_claim_entered			 = 15 'O  Claim(s) Entered
const excel_col_assigned_to				 = 16 'P' Who worker who cleared
const excel_col_match_cleared            = 17 'Q  case note to check match cleared
const excel_col_period		   		     = 18 'R' match periods
const excel_col_atr_signed				 = 19 'S' Date signed ATR
const excel_col_evf_rcvd				 = 20 'T' Date EVF Recieved
const excel_col_other_note				 = 21 'U
const excel_col_comments				 = 22 'V I dont use this


'Now the script adds all the clients on the excel list into an array
excel_row = 2 're-establishing the row to start based on when picking up the information
entry_record = 0 'incrementor for the array and count

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

Do 'purpose is to read each excel row and to add into each excel array '
 	'Reading information from the Excel
	add_to_array = FALSE
 	MAXIS_case_number = objExcel.cells(excel_row, excel_col_case_number).Value
 	MAXIS_case_number = trim(MAXIS_case_number)
	IF trim(objExcel.cells(excel_row, excel_col_period).Value) <> "" THEN add_to_array = TRUE
	IF trim(objExcel.cells(excel_row, excel_col_resolution_status).Value) <> "" THEN add_to_array = TRUE
	'MsgBox MAXIS_case_number & " " & client_SSN & " " & add_to_array
	IF MAXIS_case_number = "" THEN EXIT DO
	'msgbox "being added: " & excel_row & " " & add_to_array & vbcr & " " & MAXIS_case_number & " " & client_SSN

	IF add_to_array = TRUE THEN   'Adding client information to the array - this is for READING FROM the excel
     	ReDim Preserve match_based_array(other_note_const, entry_record)	'This resizes the array based on the number of cases
	    match_based_array(maxis_case_number_const,  entry_record)	 = MAXIS_case_number
		match_based_array(client_ssn_const, 		entry_record)	 = trim(objExcel.cells(excel_row, excel_col_client_ssn).Value)
		match_based_array(client_ssn_const, 		entry_record)	 = replace(match_based_array(client_ssn_const, entry_record), "-", "")
	    match_based_array(assigned_to_const,  		entry_record)	 = trim(objExcel.cells(excel_row, excel_col_assigned_to).Value)
	    match_based_array(worker_number_const,  	entry_record)	 = trim(objExcel.cells(excel_row, excel_col_worker_number).Value)
	    match_based_array(client_DOB_const,  		entry_record)    = trim(objExcel.cells(excel_row, excel_col_client_DOB).Value)
	    match_based_array(relationship_const,  		entry_record)    = trim(objExcel.cells(excel_row, excel_col_relationship).Value)
	    match_based_array(case_earner_name_const,  	entry_record)    = trim(objExcel.cells(excel_row, excel_col_case_earner_name).Value)
	    match_based_array(client_name_const, 		entry_record)    = trim(objExcel.cells(excel_row, excel_col_client_name).Value)
		match_based_array(program_const,  			entry_record)    = trim(objExcel.cells(excel_row, excel_col_program).Value)
	    match_based_array(amount_const,  			entry_record)    = trim(objExcel.cells(excel_row, excel_col_amount).Value)
		match_based_array(amount_const, 			entry_record) 	 = replace(match_based_array(amount_const, entry_record), "$", "")
		match_based_array(amount_const, 			entry_record)	 = replace(match_based_array(amount_const, entry_record), ",", "")
		match_based_array(amount_const, 			entry_record) 	 = trim(match_based_array(amount_const, entry_record))
	    match_based_array(income_source_const, 		entry_record)    = trim(objExcel.cells(excel_row, excel_col_income_source).Value)
	    match_based_array(resolution_status_const,  entry_record)    = trim(objExcel.cells(excel_row, excel_col_resolution_status).Value) 'does it matter I repeat this'
		match_based_array(resolution_status_const,  entry_record)    = UCASE(objExcel.cells(excel_row, excel_col_resolution_status).Value)
		match_based_array(notice_sent_date_const,  	entry_record)    = trim(objExcel.cells(excel_row, excel_date_notice_sent).Value)
		match_based_array(date_cleared_const,  		entry_record)    = trim(objExcel.cells(excel_row, excel_col_date_cleared).Value)
		match_based_array(IEVS_period_const,  		entry_record)    = trim(objExcel.cells(excel_row, excel_col_period).Value)

		match_based_array(other_note_const,  		entry_record)    = trim(objExcel.cells(excel_row, excel_col_other_note).Value)
		match_based_array(excel_row_const, entry_record) = excel_row
		'msgbox  "?" & entry_record
	    'making space in the array for these variables, but valuing them as "" for now
      	entry_record = entry_record + 1			'This increments to the next entry in the array
      	stats_counter = stats_counter + 1 'Increment for stats counter
		excel_row = excel_row + 1
	END IF
Loop

'msgbox "*" & entry_record & vbcr & " excel row " & excel_row

'Loading of cases is complete. Reviewing the cases in the array.
'msgbox " ????" & excel_row & " " & add_to_array & vbcr & " " & MAXIS_case_number & " " & client_SSN
For item = 0 to UBound(match_based_array, 2)
	MAXIS_case_number = match_based_array(maxis_case_number_const, item)
	CALL navigate_to_MAXIS_screen("INFC" , "____")
	'CALL write_value_and_transmit(client_SSN, 3, 63)
	CALL write_value_and_transmit(match_based_array(client_ssn_const, item), 3, 63)
	CALL write_value_and_transmit("IEVP", 20, 71)

	EMReadScreen current_panel_check, 4, 2, 45
	IF current_panel_check = "INFC" THEN
		EMReadScreen MISC_error_check,  74, 24, 02
		match_based_array(match_cleared_const, item) = trim(MISC_error_check)
	End IF
	'msgbox "IEVP " & match_based_array(match_cleared_const, item)
	'------------------------------------------------------------------selecting the correct wage match
	'Setting the match type'
	IF match_type = "BNDX"   THEN numb_match_type = "A3" 'removed last digit due to wage match having two numbers doing the samething'
	IF match_type = "SDXS/I" THEN numb_match_type = "A4"
	IF match_type = "BEER"   THEN numb_match_type = "A7"
	IF match_type = "UNVI"   THEN numb_match_type = "A8"
	IF match_type = "UBEN"   THEN numb_match_type = "A6"
	IF match_type = "WAGE"   THEN numb_match_type = "A5"

	Row = 7
	DO
		EMReadScreen panel_check, 4, 2, 52
		IF panel_check = "IEVP" THEN
			EMReadScreen IEVS_period, 11, row, 47
			EMReadScreen ievp_match_type, 2, row, 41 'read the match type
			EMReadScreen days_pending, 4, row, 72'
			IF ievp_match_type = numb_match_type THEN
				'msgbox " ~ " & match_based_array(IEVS_period_const, item) & " ~ " & IEVS_period
				IF match_based_array(IEVS_period_const, item) = IEVS_period  THEN
					days_pending = ""             '?? - can this be blanked out? It's always going to be a numeric otherwise.
			    	EMReadScreen days_pending, 4, row, 72
			    	days_pending = trim(days_pending)
			    	days_pending = replace(days_pending, "(", "")
			    	days_pending = replace(days_pending, ")", "")
			    	'msgbox "pending " & days_pending & " "& IsNumeric(days_pending)
		        	IF IsNumeric(days_pending) = TRUE THEN
						match_based_array(match_cleared_const, item) = TRUE
		       	   		CALL write_value_and_transmit("U", row, 3)   'navigates to IULA
		            	EMReadScreen panel_check, 4, 02, 52
		        		msgbox panel_check
	                	'IF panel_check = "IULA" THEN
	                	'----------------------------------------------------------------------------------------------------ACTIVE PROGRAMS
		                EMReadScreen Active_Programs, 1, 6, 68 'only reading one becasue I trimmed out extra in the beginning
		                IF match_type = "WAGE" THEN
		                	EMReadScreen income_line, 44, 8, 37 'should be to the right of employer and the left of amount
		                	income_line = trim(income_line)
		                	income_amount = right(income_line, 8)
		                	IF instr(income_line, " AMOUNT: $") THEN position = InStr(income_line, " AMOUNT: $")          'sets the position at the deliminator
		                	IF instr(income_line, " AMT: $") THEN position = InStr(income_line, " AMT: $")    		      'sets the position at the deliminator
		                	income_source = Left(income_line, position)  'establishes employer as being before the deliminator
		                	income_amount = replace(income_amount, "$", "")
		                	income_amount = replace(income_amount, ",", "")
		                	income_amount = trim(income_amount)
		                END IF
		                IF match_type = "BEER" THEN
		                	EMReadScreen income_line, 44, 8, 28
		                	income_line = trim(income_line)
		                	income_amount = right(income_line, 8)
		                	IF instr(income_line, " AMOUNT: $") THEN	position = InStr(income_line, " AMOUNT: $")    	  'sets the position at the deliminator
		                	IF instr(income_line, " AMT: $") THEN position = InStr(income_line, " AMT: $")    		      'sets the position at the deliminator
		                	income_source = Left(income_line, position)  'establishes employer as being before the deliminator
		                	income_amount = replace(income_amount, "$", "")
		                	income_amount = replace(income_amount, ",", "")
		                	income_amount = trim(income_amount)
		                END IF
		                'msgbox "7" & match_based_array(amount_const,  item) & vbcr & income_amount & vbcr & match_based_array(match_cleared_const, item)
		                'This is the bigger loop to exit the loop for the excel sheet
		                'IF income_source <> match_based_array(income_source_const, item) THEN match_based_array(match_cleared_const, item) = FALSE
		                'IF active_Programs <> match_based_array(program_const,  item) THEN match_based_array(match_cleared_const, item) = FALSE
		                IF income_amount = match_based_array(amount_const,  item) THEN match_based_array(match_cleared_const, item) = TRUE
		                IF match_based_array(match_cleared_const, item) = TRUE THEN
                            'msgbox "true - exit do"
                            EXIT DO
                        End if
		        		msgbox "match based array = " & income_amount = match_based_array(amount_const,  item)
		        	    IF match_based_array(match_cleared_const, item) = FALSE THEN
							PF3 'just to leave after checking to see if we matched'
                            'msgbox "false - exit do"
							EXIT DO
						END IF
					ELSEIF match_based_array(match_cleared_const, item) = FALSE THEN
						'msgbox "my date is false and I'm exiting the do"
						PF3
						'MsgBox "did I need to PF3?"
                        exit do
					END IF
		        END IF
			END IF
		END IF
	    row = row + 1
	    'msgbox "ROW " & row
	    IF row = 17 THEN
		    PF8
		    row = 7
		END IF
	LOOP UNTIL trim(IEVS_period) = "" 'two ways to leave a loop
    'msgbox "exited loop"
    '---------------------------------------------------------------------Reading potential errors for out-of-county cases
    IF match_based_array(match_cleared_const, item) = TRUE THEN
	    '--------------------------------------------------------------------IULA
		'msgbox " I think I shall go update "
    	EMReadScreen OutOfCounty_error, 12, 24, 2
		IF OutOfCounty_error = "MATCH IS NOT" THEN match_based_array(match_cleared_const, item) = FALSE
		'possiblie PRIV read here'
     	IEVS_year = "2021"
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
    		select_quarter = "YEAR"
    	END IF
		'confirm client name  '
	   	EMReadScreen client_name, 35, 5, 24
    	client_name = trim(client_name)                     	'trimming the client name
    	IF instr(client_name, ",") THEN    						'Most cases have both last name and 1st name. This separates the two names
    	length = len(client_name)                           	'establishing the length of the variable
    	position = InStr(client_name, ",")                  	'sets the position at the deliminator (in this case the comma)
    		last_name = Left(client_name, position-1)           'establishes client last name as being before the deliminator
    		first_name = Right(client_name, length-position)    'establishes client first name as after before the deliminator
    	ELSEIF instr(first_name, " ") THEN   						'If there is a middle initial in the first name, THEN it removes it
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
    		EMReadScreen income_source, 50, 8, 37
    		income_source = trim(income_source)
    		IF instr(income_source, " AMOUNT: $") THEN'should be to the right of employer and the left of amount '
    			position = InStr(income_source, " AMOUNT: $")    		      'sets the position at the deliminator
    			income_source = Left(income_source, position)  'establishes employer as being before the deliminator
    		Elseif instr(income_source, " AMT: $") THEN 					  'establishing the length of the variable
    			position = InStr(income_source, " AMT: $")    		      'sets the position at the deliminator
    			income_source = Left(income_source, position)  'establishes employer as being before the deliminator
    		END IF
    	END IF
    	IF match_type = "BEER" THEN
    		EMReadScreen income_source, 50, 8, 28
    		income_source = trim(income_source)
    		IF instr(income_source, " AMOUNT: $") THEN
    			position = InStr(income_source, " AMOUNT: $")    		      'sets the position at the deliminator
    			income_source = Left(income_source, position)  'establishes employer as being before the deliminator
    		Elseif instr(income_source, " AMT: $") THEN 					  'establishing the length of the variable
    			position = InStr(income_source, " AMT: $")    		      'sets the position at the deliminator
    			income_source = Left(income_source, position)  'establishes employer as being before the deliminator
    		END IF
    	END IF
		'This is the bigger loop to exit the loop for the excel sheet
		'----------------------------------------------------------------------------------------------------RESOLVING THE MATCH
    	EMReadScreen match_based_array(notice_sent_const,   item), 1, 14, 37
		IF match_based_array(notice_sent_const,   item) = "Y" THEN
			EMReadScreen match_based_array(notice_sent_date_const,   item), 8, 14, 68
			match_based_array(notice_sent_date_const,   item) = replace(match_based_array(notice_sent_date_const,   item), " ", "/")
		END IF
    	EMReadScreen clear_code, 2, 12, 58
		IF clear_code <> "__" THEN match_based_array(match_cleared_const, item) = FALSE 'default to false unless something happens to make it not'
		EMwriteScreen "10", 12, 46	    'resolved notes depending on the resolution_status
	   	EMwritescreen match_based_array(resolution_status_const,  item), 12, 58
		' msgbox "wrote the resolution"
		TRANSMIT 'Going to IULB
	' 	'----------------------------------------------------------------------------------------writing the note on IULB
	'
		IF match_based_array(resolution_status_const,  item) = "CB" THEN IULB_notes = "CB-Ovrpmt And Future Save"
		IF match_based_array(resolution_status_const,  item) = "CC" THEN IULB_notes = "CC-Overpayment Only"
		IF match_based_array(resolution_status_const,  item) = "CF" THEN IULB_notes = "CF-Future Save"
		IF match_based_array(resolution_status_const,  item) = "CA" THEN IULB_notes = "CA-Excess Assets"
		IF match_based_array(resolution_status_const,  item) = "CI" THEN IULB_notes = "CI-Benefit Increase"
		IF match_based_array(resolution_status_const,  item) = "CP" THEN IULB_notes = "CP-Applicant Only Savings"
		IF match_based_array(resolution_status_const,  item) = "BC" THEN IULB_notes = "BC-Case Closed"
		'IF match_based_array(resolution_status_const,  item) = "BE" THEN IULB_notes = "BE-Child" "TODO change the code above to match"
		IF match_based_array(resolution_status_const,  item) = "BE" THEN IULB_notes = "BE-No Change"
		'IF match_based_array(resolution_status_const,  item) = "BE" THEN IULB_notes = "BE-Overpayment Entered"
		'IF match_based_array(resolution_status_const,  item) = "BE" THEN IULB_notes = "BE-NC-Non-collectible"
		IF match_based_array(resolution_status_const,  item) = "BI" THEN IULB_notes = "BI-Interface Prob"
		IF match_based_array(resolution_status_const,  item) = "BN" THEN IULB_notes = "BN-Already Known-No Savings"
		IF match_based_array(resolution_status_const,  item) = "BP" THEN IULB_notes = "BP-Wrong Person"
		IF match_based_array(resolution_status_const,  item) = "BU" THEN IULB_notes = "BU-Unable To Verify"
		IF match_based_array(resolution_status_const,  item) = "BO" THEN IULB_notes = "BO-Other"
		IF match_based_array(resolution_status_const,  item) = "NC" THEN IULB_notes = "NC-Non Cooperation"

		EMReadScreen panel_name, 4, 02, 52
	    IF panel_name = "IULB" THEN
	  		EMWriteScreen IULB_notes, 8, 6
			'msgbox "we wrote the note"
	    	EMReadScreen IULB_enter_msg, 5, 24, 02
	    	IF IULB_enter_msg = "ENTER" OR IULB_enter_msg = "ACTIO" THEN 'check if we need to input other notes
				CALL clear_line_of_text(8, 6)
				CALL clear_line_of_text(9, 6)
			END IF
			IULB_comment = ""
			IF IULB_notes = "CA-Excess Assets" THEN IULB_comment = "Excess Assets. " & other_notes
			IF IULB_notes = "CI-Benefit Increase" THEN IULB_comment = "Benefit Increase. " & other_notes
			IF IULB_notes = "CP-Applicant Only Savings" THEN IULB_comment = "Applicant Only Savings. " & other_notes
			IF IULB_notes = "BC-Case Closed" THEN IULB_comment = "Case closed. " & other_notes
			'IF IULB_notes = "BE-Child" THEN IULB_comment = "No change, minor child income excluded. " & other_notes
			IF IULB_notes = "BE-No Change" THEN IULB_comment = "No change. " & other_notes
			'IF IULB_notes = "BE-Overpayment Entered" THEN IULB_comment = "OP entered other programs. " & other_notes
			'IF IULB_notes = "BE-NC-Non-collectible" THEN IULB_comment = "Non-Coop remains, but claim is non-collectible. " & other_notes
			IF IULB_notes = "BI-Interface Prob" THEN IULB_comment = "Interface Problem. " & other_notes
			IF IULB_notes = "BN-Already Known-No Savings" THEN IULB_comment = "Already known - No savings. " & other_notes
			IF IULB_notes = "BP-Wrong Person" THEN IULB_comment = "Client name and wage earner name are different. " & other_notes
			IF IULB_notes = "BU-Unable To Verify" THEN IULB_comment = "Unable To Verify. " & other_notes
			IF IULB_notes = "NC-Non Cooperation" THEN IULB_comment = "Non-coop, requested verf not in ECF, " & other_notes

			IULB_comment = trim(IULB_comment)
			iulb_row = 8
			iulb_col = 6
			notes_array = split(IULB_comment, " ")
			For each word in notes_array
				EMWriteScreen word & " ", iulb_row, iulb_col
			 	'msgbox "Word - " & word & vbCr & "Row - " & iulb_row & "   Col - " & iulb_col & vbCr & "Add - " & iulb_col + len(word)
				If iulb_col + len(word) > 77 Then
					iulb_col = 6
					iulb_row = iulb_row + 1
					If iulb_row = 10 Then Exit For
				End If
				iulb_col = iulb_col + len(word) + 1
			Next
			' msgbox "NOTE WROTE"

	    	TRANSMIT
			'------------------------------------------------------------------back on the IEVP menu, making sure that the match cleared
			EMReadScreen days_pending, 5, row, 72
	    	days_pending = trim(days_pending)
	    	IF IsNumeric(days_pending) = TRUE THEN match_based_array(match_cleared_const, item) = FALSE
			' MsgBox "Cleared? " & match_based_array(match_cleared_const, item)
			'msgbox "Fini"
   		END IF
	 	    '------------------------------------------------------------------STAT/MISC for claim referral tracking
   	    IF claim_referral_tracking_dropdown <> "Not Needed" THEN
	    'Going to the MISC panel to add claim referral tracking information
	    	CALL navigate_to_MAXIS_screen ("STAT", "MISC")
	       	Row = 6
	       	EMReadScreen panel_number, 1, 02, 73
	       	IF panel_number = "0" THEN
	    	   EMWriteScreen "NN", 20,79
	    	ELSE
   	    		Do
	       			'Checking to see if the MISC panel is empty, if not it will find a new line'
	       			EMReadScreen MISC_description, 25, row, 30
	       			MISC_description = replace(MISC_description, "_", "")
	       			IF trim(MISC_description) = "" THEN
	    	   			'PF9
	    	   			EXIT DO
	       			ELSE
	    	   			row = row + 1
	       			END IF
   	    		Loop Until row = 17
   	    		If row = 17 THEN MsgBox("There is not a blank field in the MISC panel. Please delete a line(s), and run script again or update manually.")
	    	END IF
        	'writing in the action taken and date to the MISC panel
        	PF9
        	'_________________________ 25 characters to write on MISC
        	IF claim_referral_tracking_dropdown =  "Initial" THEN MISC_action_taken = "Claim Referral Initial"
        	IF claim_referral_tracking_dropdown =  "OP Non-Collectible (please specify)" THEN MISC_action_taken = "Determination-Non-Collect"
        	IF claim_referral_tracking_dropdown =  "No Savings/Overpayment" THEN MISC_action_taken = "Determination-No Savings"
        	IF claim_referral_tracking_dropdown =  "Overpayment Exists" THEN MISC_action_taken =  "Determination-OP Entered" '"Claim Determination 25 character available
        	EMWriteScreen MISC_action_taken, Row, 30
        	EMWriteScreen date, Row, 66
        	TRANSMIT
            start_a_blank_case_note
            IF claim_referral_tracking_dropdown =  "Initial" THEN CALL write_variable_in_case_note("Claim Referral Tracking - Initial")
            CALL write_bullet_and_variable_in_case_note("Action Date", Date)
            CALL write_bullet_and_variable_in_case_note("Active Program(s)", programs)
            CALL write_bullet_and_variable_in_case_note("Other Notes", other_notes)
            CALL write_variable_in_case_note("* Entries for these potential claims must be retained until further notice.")
            IF case_note_only = TRUE THEN CALL write_variable_in_case_note("Maxis case is inactive unable to add or update MISC panel")
            CALL write_variable_in_case_note("-----")
            CALL write_variable_in_case_note(worker_signature)
            PF3
        END IF

    	'-------------------------------------------------------------------------------------------------for the case note
    	IF match_type = "BEER" THEN match_type_letter = "B"
        IF match_type = "UBEN" THEN match_type_letter = "U"
        IF match_type = "UNVI" THEN match_type_letter = "U"

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


		MsgBox "Start of case note what screen"
		'IEVS_quarter = "2ND"
		'assignment_date = "08/18/21"
		'worker_number = "X127D5X"

        CALL navigate_to_MAXIS_screen_review_PRIV("CASE", "NOTE", is_this_priv)
		EMReadScreen county_code, 4, 21, 14  'Out of county cases from STAT
		EMReadScreen case_invalid_error, 72, 24, 2 'if a person enters an invalid footer month for the case the script will attempt to  navigate'
		IF priv_check = TRUE THEN  'PRIV cases
			EMReadscreen priv_worker, 26, 24, 46
			match_based_array(other_note_const, item) = trim(priv_worker)
			match_based_array(match_cleared_const, item) = FALSE
			' MsgBox "We think it is PRIV"
		ELSEIf county_code <> "X127" THEN
		  match_based_array(other_note_const, item) = "OUT OF COUNTY CASE"
		  match_based_array(match_cleared_const, item) = FALSE
		  ' MsgBox "We think it is Out of County"
		ELSEIF instr(case_invalid_error, "IS INVALID") THEN  'CASE xxxxxxxx IS INVALID FOR PERIOD 12/99
			match_based_array(other_note_const, item) = trim(case_invalid_error)
			match_based_array(match_cleared_const, item) = FALSE
			' MsgBox "INVALID?"
		ELSE
			EMReadScreen MAXIS_case_name, 27, 21, 40 'not always the same as the match name'
			MAXIS_row = 6
			DO
				EMReadscreen case_note_date, 8, MAXIS_row, 6

				IF trim(case_note_date) = "" THEN
					match_based_array(other_note_const, item) = "NO CASE NOTE"
					match_based_array(match_cleared_const, item) = FALSE
					' MsgBox "NO NOTE??"
					EXIT DO
				ELSE
					IF case_note_date = assignment_date THEN 'weekends and the day prior has the date assigned confirmed by the SSR '
						match_based_array(other_note_const, item) = case_note_date
						EMReadScreen case_note_worker_number, 7, MAXIS_row, 16

						MsgBox worker_number & "~" & case_note_worker_number & VBCR & assignment_date & " ~ " & case_note_date

						IF worker_number = case_note_worker_number THEN
							'IF worker_number = "X127D5X" THEN match_based_array(match_cleared_const, item) = FALSE
							'IF worker_number = "X127823" THEN match_based_array(match_cleared_const, item) = FALSE
							match_based_array(other_note_const, item) = "CASE NOTED"
							EMReadScreen case_note_header, 55, MAXIS_row, 25
							case_note_header = lcase(trim(case_note_header))
							IF instr(case_note_header, "wage match") then match_based_array(other_note_const, item) = "DUPLICATE"
							EXIT DO
						END IF
					END IF
				END IF
			    MAXIS_row = MAXIS_row + 1
			    IF MAXIS_row = 19 THEN
			    	PF8 'moving to next case note page if at the end of the page
			    	MAXIS_row = 5
			    END IF
			LOOP UNTIL case_note_date => cdate(assignment_date)   'repeats until the case note date is less than the assignment date

		    programs = ""
		    IF instr(match_based_array(program_const,  			item) , "D") THEN programs = programs & "DWP, "
		    IF instr(match_based_array(program_const,  			item) , "F") THEN programs = programs & "Food Support, "
		    IF instr(match_based_array(program_const,  			item) , "H") THEN programs = programs & "Health Care, "
		    IF instr(match_based_array(program_const,  			item) , "M") THEN programs = programs & "Medical Assistance, "
		    IF instr(match_based_array(program_const,  			item) , "S") THEN programs = programs & "MFIP, "
		    'trims excess spaces of programs
		    programs = trim(programs)
		    'takes the last comma off of programs when autofilled into dialog
		    IF right(programs, 1) = "," THEN programs = left(programs, len(programs) - 1)
			IF match_based_array(match_cleared_const, item) = TRUE THEN
    		    PF9
        	    IF match_type = "WAGE" THEN CALL write_variable_in_case_note("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH"  & " (" & first_name & ") CLEARED " & match_based_array(resolution_status_const,  item) & "-----")
        	    IF match_type = "BEER" THEN CALL write_variable_in_case_note("-----" & IEVS_year & " NON-WAGE MATCH(" & match_type_letter & ")" & " (" & first_name & ") CLEARED " & match_based_array(resolution_status_const,  item) & "-----")
        	    IF match_type = "UNVI" THEN CALL write_variable_in_case_note("-----" & IEVS_year & " NON-WAGE MATCH(" & match_type_letter & ")" & " (" & first_name & ") CLEARED " & match_based_array(resolution_status_const,  item) & "-----")
        	    IF match_type = "UBEN" THEN CALL write_variable_in_case_note("-----" & IEVS_period & " NON-WAGE MATCH(" & match_type_letter & ")" & " (" & first_name & ") CLEARED " & match_based_array(resolution_status_const,  item) & "-----")
        	    IF match_type = "BNDX" THEN CALL write_variable_in_case_note("-----" & IEVS_period & " NON-WAGE MATCH(" & match_type & ")" & " (" & first_name & ") CLEARED " & match_based_array(resolution_status_const,  item) & "-----")
        	    CALL write_bullet_and_variable_in_case_note("Period", match_based_array(IEVS_period_const, item))
        	    CALL write_bullet_and_variable_in_case_note("Active Programs", programs)
        	    CALL write_bullet_and_variable_in_case_note("Source of income", match_based_array(income_source_const, item))
        	    CALL write_bullet_and_variable_in_case_note("Date Diff notice sent", match_based_array(notice_sent_date_const, item))
        	    IF IULB_notes = "CB-Ovrpmt And Future Save" THEN CALL write_variable_in_case_note("* OP Claim entered and future savings.")
        	    IF IULB_notes = "CF-Future Save" THEN CALL write_variable_in_case_note("* Future Savings.")
        	    IF IULB_notes = "CA-Excess Assets" THEN CALL write_variable_in_case_note("* Excess Assets.")
        	    IF IULB_notes = "CI-Benefit Increase" THEN CALL write_variable_in_case_note("* Benefit Increase.")
        	    IF IULB_notes = "CP-Applicant Only Savings" THEN CALL write_variable_in_case_note("* Applicant Only Savings.")
        	    IF IULB_notes = "BC-Case Closed" THEN CALL write_variable_in_case_note("* Case closed.")
        	    IF IULB_notes = "BE-Child" THEN CALL write_variable_in_case_note("* Income is excluded for minor child in school.")
        	    IF IULB_notes = "BE-No Change" THEN CALL write_variable_in_case_note("* No Overpayments or savings were found related to this match.")
        	    IF IULB_notes = "BE-Overpayment Entered" THEN CALL write_variable_in_case_note("* Overpayments or savings were found related to this match.")
        	    IF IULB_notes = "BE-NC-Non-collectible" THEN CALL write_variable_in_case_note("* No collectible overpayments or savings were found related to this match. Client is still non-coop.")
        	    IF IULB_notes = "BI-Interface Prob" THEN CALL write_variable_in_case_note("* Interface Problem.")
        	    IF IULB_notes = "BN-Already Known-No Savings" THEN CALL write_variable_in_case_note("* Client reported income. Correct income is in JOBS/BUSI and budgeted.")
        	    IF IULB_notes = "BP-Wrong Person" THEN CALL write_variable_in_case_note("* Client name and wage earner name are different.  Client's SSN has been verified. No overpayment or savings related to this match.")
        	    IF IULB_notes = "BU-Unable To Verify" THEN CALL write_variable_in_case_note("* Unable to verify, due to:")
        	    IF IULB_notes = "BO-Other" THEN CALL write_variable_in_case_note("* No review due during the match period.  Per DHS, reporting requirements are waived during pandemic.")
        	    CALL write_bullet_and_variable_in_case_note("Other Notes", other_notes)
        	    CALL write_variable_in_case_note("----- ----- ----- ----- -----")
        	    CALL write_variable_in_case_note("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1")
        	    PF3 'to save casenote'
        		match_based_array(match_cleared_const, item) = TRUE
			END IF
        END IF
 	END IF
NEXT

'Excel headers and formatting the columns
objExcel.Cells(1, 1).Value     = "DATE POSTED" 		'A' Date Posted to Maxis'
objExcel.Cells(1, 2).Value     = "BASKET" 			'B' Worker #
objExcel.Cells(1, 3).Value     = "DOB" 				'C' DOB
objExcel.Cells(1, 4).Value     = "RELATIONSHIP" 	'D' Relationship
objExcel.Cells(1, 5).Value     = "EARNER NAME" 		'E' Earner Name
objExcel.Cells(1, 6).Value     = "CASE NUMBER" 		'F' Case Number
objExcel.Cells(1, 7).Value     = "CLIENT NAME" 		'G' Case Name
objExcel.Cells(1, 8).Value     = "SSN" 				'H' SSN
objExcel.Cells(1, 9).Value     = "PROG"				'I' Prog
objExcel.Cells(1, 10).Value    = "AMOUNT"			'J' Amount
objExcel.Cells(1, 11).Value    = "SOURCE OF INCOME" 'K' Employer
objExcel.Cells(1, 12).Value    = "NOTICE SENT"		'L' Date Notice Sent
objExcel.Cells(1, 13).Value    = "RESOLUTION"		'M' How cleared
objExcel.Cells(1, 14).Value    = "DATE CLEARED"		'N
objExcel.Cells(1, 15).Value    = "DATE CLAIM ENTERED"' Claim(s) Entered
objExcel.Cells(1, 16).Value    = "ASSIGNED TO"		'P' Who worker who cleared
objExcel.Cells(1, 17).Value    = "MATCH CLEARED"	'Q  case note to check match cleared/used to be work# requested
objExcel.Cells(1, 18).Value    = "PERIOD"			'R' Date cleared
objExcel.Cells(1, 19).Value    = "DATE ATR SIGNED"	'S  Date signed ATR
objExcel.Cells(1, 20).Value    = "DATE EVF RCVD"	'T  Date EVF Recieved
objExcel.Cells(1, 21).Value    = "OTHER NOTES"		'U  OTHER NOTES
objExcel.Cells(1, 22).Value    = "COMMENTS"			'V  OTHER NOTES

For item = 0 to UBound(match_based_array, 2)
 	excel_row = match_based_array(excel_row_const, item)
 	objExcel.Cells(excel_row, excel_col_match_cleared).Value 	= match_based_array(match_cleared_const, item)
	objExcel.Cells(excel_row, excel_col_other_note).Value 		= match_based_array(other_note_const,   item)
	objExcel.Cells(excel_row, excel_date_notice_sent).Value		= match_based_array(notice_sent_date_const,   item)
	IF match_based_array(match_cleared_const, item) = TRUE THEN objExcel.Cells(excel_row, excel_col_date_cleared).Value = date
Next

FOR i = 1 to 23		'formatting the cells
    objExcel.Cells(1, i).Font.Bold 			= TRUE		'bold font'
    objExcel.Columns(i).AutoFit()						'sizing the columns'
	objExcel.Columns(i).HorizontalAlignment = -4131		'Centers the text for the columns
	objExcel.Columns(12).NumberFormat = "mm/dd/yy"		'Date Notice Sent
	objExcel.Columns(14).NumberFormat = "mm/dd/yy"		'Date Cleared
	objExcel.Columns(15).NumberFormat = "mm/dd/yy"		'Date Notice Sent
NEXT

STATS_counter = STATS_counter - 1   'subtracts one from the stats (since 1 was the count, -1 so it's accurate)

script_end_procedure_with_error_report("Success your list has been updated, please review to ensure accuracy.")
