'Required for statistical purposes==========================================================================================
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

Function IEVP_looping(ievp_panel)
    row = row + 1
    IF row = 17 THEN
        PF8
        row = 7
        EMReadScreen IVEP_panel_check, 4, 2, 52
        IF IEVP_panel_check = "IEVP" THEN
            IEVP_panel = True
        Else
            EMReadScreen MISC_error_check,  74, 24, 02
            match_based_array(comments_const, item) = "Unable to navigate to IEVP. Error message: " & trim(MISC_error_check)
            'msgbox "1. IN IEVP LOOPING" & trim(MISC_error_check)
            IEVP_panel = False
        End IF
    End if
End Function

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("01/26/2023", "Removed term 'ECF' from the case note per DHS guidance, and referencing the case file instead.", "Ilse Ferris, Hennepin County")
call changelog_update("11/15/2022", "Resolved bug in excel row incrementor which was previously skipping cases.", "Ilse Ferris, Hennepin County")
CALL changelog_update("10/06/2022", "Update to remove hard coded DEU signature all DEU scripts.", "MiKayla Handley, Hennepin County") '#316
CALL changelog_update("09/16/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
CALL changelog_update("07/26/2022", "Updated handling for claim referral tracking.", "MiKayla Handley, Hennepin County") '#991
CALL changelog_update("06/30/2022", "Updated handling for new upervisor.", "MiKayla Handley, Hennepin County") '#498
CALL changelog_update("06/21/2022", "Updated handling for non-disclosure agreement and closing documentation.", "MiKayla Handley, Hennepin County") '#493
call changelog_update("08/19/2021", "Retire the BULK script due to redundancy.", "MiKayla Handley, Hennepin County") '#596'
call changelog_update("06/17/2021", "Updating the dialog box to ensure that a cleared method is entered.", "MiKayla Handley, Hennepin County") '#498'
call changelog_update("12/07/2019", "Added handling for coding the Excel spreadsheet. You must use BC, BE, BN, or CC only in the cleared status field.", "MiKayla Handley, Hennepin County")
call changelog_update("11/14/2017", "Program information will not be input into the Excel spreadsheet. This will not need to be added manually by staff completing the cases.", "Ilse Ferris, Hennepin County")
call changelog_update("06/05/2017", "Added handling for minor children in school (excluded income) & multiple people per case.", "Ilse Ferris, Hennepin County")
call changelog_update("03/20/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr
match_type = "WAGE"

'This can only be run by DEU Supervisor or script team member - Victoria, Ilse or Casey
IF user_ID_for_validation <> "WFO119" THEN
    IF user_ID_for_validation <> "ILFE001" THEN
		IF user_ID_for_validation <> "CALO001" THEN
	   		script_end_procedure("This is restricted to use by a supervisor only. Please contact your supervisor to run.")
		END IF
	END IF
END IF

Do
	Do
		'The dialog is defined in the loop as it can change as buttons are pressed
	    '-------------------------------------------------------------------------------------------------DIALOG
	    Dialog1 = "" 'Blanking out previous dialog detail
        BeginDialog Dialog1, 0, 0, 271, 240, "BULK-Match Cleared"
          DropListBox 140, 15, 120, 15, "Select One:"+chr(9)+"BEER"+chr(9)+"BNDX"+chr(9)+"SDXS/SDXI"+chr(9)+"UNVI"+chr(9)+"UBEN"+chr(9)+"WAGE", match_type
          EditBox 65, 35, 195, 15, other_notes
          ButtonGroup ButtonPressed
            PushButton 10, 115, 50, 15, "Browse:", select_a_file_button
          EditBox 65, 115, 195, 15, IEVS_match_path
          ButtonGroup ButtonPressed
            PushButton 60, 175, 145, 15, "Open IEVS Template Excel File", open_ievs_template_file_button
          EditBox 70, 220, 95, 15, worker_signature
          ButtonGroup ButtonPressed
            OkButton 170, 220, 45, 15
            CancelButton 220, 220, 45, 15
          GroupBox 5, 5, 260, 65, "Complete prior to browsing the script:"
          Text 10, 20, 120, 10, "Select the type of match to process:"
          Text 10, 40, 45, 10, "Other Notes:"
          Text 10, 55, 170, 10, "Please note this will apprear on ALL cases cleared."
          GroupBox 5, 75, 260, 140, "Using the script:"
          Text 10, 90, 250, 15, "Select the Excel file that contains the case information by selecting the 'Browse' button and locating the file."
          Text 10, 135, 245, 15, "This script should be used when matches have been researched and ready to be cleared. "
          Text 10, 155, 245, 20, "You MUST use the correct Excel layout for this script to work properly. The column positions and layout can be found in the IEVS Template Excel file."
          Text 10, 195, 245, 20, "If you use a different layout in the file you select, the script will likely not function correctly."
          Text 5, 225, 60, 10, "Worker Signature:"
        EndDialog

	  	err_msg = ""
		DIALOG Dialog1
		cancel_confirmation
		If ButtonPressed = select_a_file_button then call file_selection_system_dialog(IEVS_match_path, ".xlsx")
		If match_type = "Select One:" then err_msg = err_msg & vbNewLine & "* Please select type of match you are processing."
        If trim(worker_signature) = "" then err_msg = err_msg & vbNewLine & "* Please enter your worker signature."
 		If ButtonPressed = open_ievs_template_file_button Then
			err_msg = "LOOP"
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/BlueZone%20Script%20Resources/IEVS%20TEMPLATE.xlsx"
		End If
		If err_msg <> "" and err_msg <> "LOOP" Then MsgBox err_msg
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = FALSE					'loops until user passwords back in

Call excel_open(IEVS_match_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file

'setting the footer month to make the updates in'
back_to_self 'resetting MAXIS back to self before getting started
Call MAXIS_footer_month_confirmation	'ensuring we are in the correct footer month/year

'setting the columns - using constant so that we know what is going on'
'const excel_col_date_posted_to_maxis	 = 1 'A' 'Date Posted to Maxis'
'const excel_col_worker_number 			 = 2 'B' Worker #
'const excel_col_client_DOB				 = 3 'C' DOB
'const excel_col_relationship			 = 4 'D' Relationship
const excel_col_case_number   			 = 5 'E' Case Number
const excel_col_case_earner_name 	     = 6 'F' Earner Name
const excel_col_client_name				 = 7 'G' Case Name
const excel_col_client_ssn				 = 8 'H' SSN
const excel_col_program  				 = 9 'I' Program
const excel_col_amount 					 = 10 'J' Amount
const excel_col_income_source		     = 11 'K' Employer
const excel_date_notice_sent		     = 12 'L' Date Notice Sent
const excel_col_resolution_status   	 = 13 'M' How cleared
const excel_col_date_cleared			 = 14 'N' Date cleared
const excel_col_claim_entered			 = 15 'O  Claim(s) Entered
const excel_col_assigned_to				 = 16 'P' Assigned to
const excel_col_numb_match_type          = 17 'Q  MatchType
const excel_col_period		   		     = 18 'R' match periods
const excel_col_atr_signed				 = 19 'S' Date signed ATR
const excel_col_evf_rcvd				 = 20 'T' Date EVF Received
const excel_col_other_note				 = 21 'U Other Notes
const excel_col_comments				 = 22 'V Comments

'Now the script adds all the clients on the excel list into an array
excel_row = 2 're-establishing the row to start based on when picking up the information
entry_record = 0 'incrementor for the array and count

'Establishing array
DIM match_based_array()  'Declaring the array this is what this list is
ReDim match_based_array(comments_const, 0)  'Resizing the array 'that ,list is going to have 20 parameter but to start with there is only one paparmeter it gets complicated - grid'
'Creating constants to value the array elements this is why we create constants
'for each row the column is going to be the same information type
const date_posted_to_maxis_const	 		= 0 'Date Posted to Maxis
const worker_number_const			    	= 1 'Basket
const client_DOB_const 				     	= 2 'DOB
const relationship_const			     	= 3 'Relationship
const maxis_case_number_const      			= 4 'Case #
const case_earner_name_const	   			= 5 'Earner Name
const client_name_const				     	= 6 'Case Name
const client_ssn_const				    	= 7 'SSN
const program_const  				       	= 8 'Prog
const amount_const 					 	    = 9 'Amount
const income_source_const		     	 	= 10 'Employer
const notice_sent_const		     		  	= 11 'Notice Sent y/n
const notice_sent_date_const		    	= 12 'Date Notice Sent
const resolution_status_const   	 		= 13 'How cleared
const date_cleared_const			 	    = 14 'Date cleared
const claim_entered_const			 	    = 15 'Claim entered
const assigned_to_const				 	    = 16 'Assigned to
const numb_match_type_const					= 17 'Match Type
const period_const	                        = 18 'Match periods
const atr_signed_const	                    = 19 'Date ATR on file
const evf_rcvd_const	                    = 20 'Date EVF Received
const priv_case_const      					= 21
const out_of_county_const 					= 22
const other_notes_const	                    = 23 'other notes
const match_cleared_const				    = 24 'true/false
const comments_const	                    = 25 'Comments

'dialog and dialog DO...Loop
Do 'purpose is to read each excel row and to add into each excel array '
 	'Reading information from the Excel
	add_to_array = FALSE
	MAXIS_case_number = objExcel.cells(excel_row, excel_col_case_number).Value
	MAXIS_case_number = trim(MAXIS_case_number)
    IF MAXIS_case_number = "" THEN EXIT DO
	IF trim(objExcel.cells(excel_row, excel_col_period).Value) <> "" THEN
        IF trim(objExcel.cells(excel_row, excel_col_resolution_status).Value) = "" THEN
			match_based_array(comments_const, item) = "No resolution status could be found."
	    ELSE
		    add_to_array = TRUE
     	    ReDim Preserve match_based_array(comments_const, entry_record)	'This resizes the array based on the number of cases
	   	    match_based_array(maxis_case_number_const,  entry_record)	 = MAXIS_case_number
	   	    match_based_array(client_ssn_const, 		entry_record)	 = trim(replace(objExcel.cells(excel_row, excel_col_client_ssn), "-", ""))
		    match_based_array(program_const,  			entry_record)    = trim(objExcel.cells(excel_row, excel_col_program).Value)
		    match_based_array(amount_const, 			entry_record) 	 = trim(objExcel.cells(excel_row, excel_col_amount).Value)
            match_based_array(amount_const, 		    entry_record)    = replace(match_based_array(amount_const, entry_record), "$", "")
            match_based_array(amount_const, 		    entry_record)    = replace(match_based_array(amount_const, entry_record), ",", "")
            match_based_array(amount_const, 		    entry_record)    = FormatNumber(match_based_array(amount_const, entry_record), 2, 0, 0, 0) 'this is formating to help the script read the number as a number'
            match_based_array(amount_const, 		    entry_record)    = match_based_array(amount_const, entry_record) *1 'this is so the amount wil be read as a number'
       	    match_based_array(income_source_const, 		entry_record)    = trim(objExcel.cells(excel_row, excel_col_income_source).Value)
	   	    match_based_array(notice_sent_date_const,  	entry_record)    = trim(objExcel.cells(excel_row, excel_date_notice_sent).Value)
	   	    match_based_array(resolution_status_const,  entry_record)    = trim(UCASE(objExcel.cells(excel_row, excel_col_resolution_status).Value)) 'does it matter I repeat this'
	   	    match_based_array(date_cleared_const,       entry_record)    = trim(objExcel.cells(excel_row, excel_col_date_cleared).Value)	' = 14 'N' Date cleared
	   	    match_based_array(numb_match_type_const,    entry_record)    = trim(objExcel.cells(excel_row, excel_col_numb_match_type).Value)   ' = 17 'Q  case note to check match cleared
	   	    match_based_array(period_const, 			entry_record)	 = replace(objExcel.cells(excel_row, excel_col_period).Value, "-", "/") ' the format that the excel sheet has is 10/21-12/21 maxis has
		    match_based_array(match_cleared_const,      entry_record)    = False    'Defaulting to false
		    match_based_array(other_notes_const,  		entry_record)    = trim(objExcel.cells(excel_row, excel_col_other_note).Value)
		    match_based_array(excel_row_const, entry_record) = excel_row
      	    entry_record = entry_record + 1			'This increments to the next entry in the array
        End If
	END IF
    excel_row = excel_row + 1
Loop

If entry_record = 0 then script_end_procedure_with_error_report("No information could be found on the Excel template. Please review/update the Excel workbook and run the script again. The script will now end.")

'Loading of cases is complete. Reviewing the cases in the array.
For item = 0 to UBound(match_based_array, 2)
	MAXIS_case_number = match_based_array(maxis_case_number_const, item)
	CALL navigate_to_MAXIS_screen("INFC" , "____")
	CALL write_value_and_transmit(match_based_array(client_ssn_const, item), 3, 63)
	CALL write_value_and_transmit("IEVP", 20, 71) 'this comes after to avoid moving away from IEVP'
	'checking for NON-DISCLOSURE AGREEMENT REQUIRED FOR ACCESS TO IEVS FUNCTIONS'
	EMReadScreen agreement_check, 9, 2, 24
	IF agreement_check = "Automated" THEN script_end_procedure("To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the screens and review the agreement.")

    EMReadScreen IEVP_panel_check, 4, 2, 52
	IF IEVP_panel_check = "IEVP" THEN
	'------------------------------------------------------------------selecting the correct wage match
        pending_match_found = False 'defaulting to FALSE
        Row = 7
	    DO
	    	EMReadScreen IEVS_period, 11, row, 47
			IEVS_period = trim(IEVS_period)
	    	IEVS_period = replace(IEVS_period, "-", "/")
		   	EMReadScreen ievp_match_type, 3, row, 41 'read the match type
            ievp_match_type = trim(ievp_match_type)

			IF ievp_match_type = "A30" THEN match_type = "BNDX"
			IF ievp_match_type = "A40" THEN match_type = "SDXS/I"
			IF ievp_match_type = "A70" THEN match_type = "BEER"
			IF ievp_match_type = "A80" THEN match_type = "UNVI"
			IF ievp_match_type = "A60" THEN match_type = "UBEN"
			IF ievp_match_type = "A50" THEN match_type = "WAGE"
			IF ievp_match_type = "A51" THEN match_type = "WAGE"
			IEVS_year = ""

			EMReadScreen days_pending, 5, row, 72
		    days_pending = trim(days_pending)
			days_pending = replace(days_pending, "(", "")
			days_pending = replace(days_pending, ")", "")
			IF IsNumeric(days_pending) = TRUE THEN
                If ievp_match_type = "" THEN
                    match_based_array(comments_const, item) = "Unable to match the IEVS types."
                    'msgbox "2. Unable to match the IEVS types."
                    exit do
                Elseif ievp_match_type = match_based_array(numb_match_type_const, item) THEN
	    			IF trim(match_based_array(period_const, item)) = IEVS_period THEN
                        pending_match_found = True
                    	CALL write_value_and_transmit("U", row, 3)   'navigates to IULA
						'----------------------------------------------------------------------------------------------------Employer info & difference notice info
						IF match_type = "UBEN" THEN income_source = "Unemployment"
						IF match_type = "UNVI" THEN income_source = "NON-WAGE"
	                    IF match_type = "WAGE" THEN	EMReadScreen income_line, 44, 8, 37 'should be to the right of employer and the left of amount
	                 	IF match_type = "BEER" THEN EMReadScreen income_line, 44, 8, 28

	                    income_line = trim(income_line)
	                    income_amount = right(income_line, 8)
	                    IF instr(income_line, " AMOUNT: $") THEN position = InStr(income_line, " AMOUNT: $")    	  'sets the position at the deliminator
	                    IF instr(income_line, " AMT: $") THEN position = InStr(income_line, " AMT: $")    		      'sets the position at the deliminator
						income_line = trim(income_line)
						income_source = Left(income_line, position)  'establishes employer as being before the deliminator
						income_source = trim(income_source)
						income_amount = replace(income_amount, "$", "")
	                    income_amount = replace(income_amount, ",", "")
						income_amount = trim(income_amount)
                        income_amount = income_amount *1 'this is so the amount wil be read as a number'

	                   	IF income_source = match_based_array(income_source_const, item) THEN
                        	IF income_amount = match_based_array(amount_const, item) THEN
							   	EXIT DO
	    				   	ELSE
							  	match_based_array(comments_const, item) = "Match not cleared due to income information" & " ~" & income_amount & "~" & match_based_array(amount_const, item) & "~"
                                'msgbox "3. Match not cleared due to income information" & " ~" & income_amount & "~" & match_based_array(amount_const, item) & "~"
                            	PF3 ' to leave match
							  	EXIT DO
							END IF
                        Else
							match_based_array(comments_const, item) = "Match not cleared due to income name information" & " ~" & income_source & "~" & match_based_array(income_source_const, item) & "~"
                            'msgbox "4. Match not cleared due to income name information" & " ~" & income_source & "~" & match_based_array(income_source_const, item) & "~"
                            Call IEVP_looping(ievp_panel)
                            If IEVP_panel = False then
                                EXIT DO
                            End if
						END IF
                        Call IEVP_looping(ievp_panel)
                        If IEVP_panel = False then
                            EXIT DO
                        End if
	    			END IF
                    Call IEVP_looping(ievp_panel)
                    If IEVP_panel = False then
                        EXIT DO
                    End if
                END IF
                Call IEVP_looping(ievp_panel)
                If IEVP_panel = False then
                    EXIT DO
                End if
			END IF
            Call IEVP_looping(ievp_panel)
		LOOP UNTIL trim(IEVS_period) = "" 'two ways to leave a loop
	ELSE
		EMReadScreen MISC_error_check,  74, 24, 02
    	match_based_array(comments_const, item) = "Unable to navigate to IEVP. Error message: " & trim(MISC_error_check)
        'msgbox "6. " & trim(MISC_error_check)
	END IF

    If pending_match_found = False then match_based_array(comments_const, item) = "Unable to find a pending match to match the match type and/or the match period. Review manually."

	'--------------------------------------------------------------------clearing the match IULA much of this is just for the case note
    IF trim(match_based_array(comments_const, item)) = "" then
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
		'client name for case note  '
	   	EMReadScreen client_name, 35, 5, 24
    	client_name = trim(client_name)                     	'trimming the client name
    	IF instr(client_name, ",") THEN    						'Most cases have both last name and 1st name. This separates the two names
    		length = len(client_name)                           	'establishing the length of the variable
    		position = InStr(client_name, ",")                  	'sets the position at the deliminator (in this case the comma)
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

		'-------------------------------------------------------------------------------------------RESOLVING THE MATCH
    	EMReadScreen match_based_array(notice_sent_const,   item), 1, 14, 37
		IF match_based_array(notice_sent_const,   item) = "Y" THEN
			EMReadScreen match_based_array(notice_sent_date_const,   item), 8, 14, 68
			match_based_array(notice_sent_date_const,   item) = replace(match_based_array(notice_sent_date_const,   item), " ", "/")
		END IF
    	EMReadScreen cleared_code, 2, 12, 58
		IF cleared_code <> "__" THEN match_based_array(resolution_status_const, item) = "CLEARED - "  & cleared_code'default to false unless something happens to make it not'
		EMwriteScreen "10", 12, 46	    'resolved notes depending on the resolution_status
	   	EMwritescreen match_based_array(resolution_status_const,  item), 12, 58
		TRANSMIT 'Going to IULB
	 	'----------------------------------------------------------------------------------------writing the note on IULB
		IF match_based_array(resolution_status_const,  item) = "CB" THEN IULB_notes = "CB-Ovrpmt And Future Save"
		IF match_based_array(resolution_status_const,  item) = "CC" THEN IULB_notes = "CC-Overpayment Only"
		IF match_based_array(resolution_status_const,  item) = "CF" THEN IULB_notes = "CF-Future Save"
		IF match_based_array(resolution_status_const,  item) = "CA" THEN IULB_notes = "CA-Excess Assets"
		IF match_based_array(resolution_status_const,  item) = "CI" THEN IULB_notes = "CI-Benefit Increase"
		IF match_based_array(resolution_status_const,  item) = "CP" THEN IULB_notes = "CP-Applicant Only Savings"
		IF match_based_array(resolution_status_const,  item) = "BC" THEN IULB_notes = "BC-Case Closed"
		IF match_based_array(resolution_status_const,  item) = "BE" THEN IULB_notes = "BE-No Change"
		IF match_based_array(resolution_status_const,  item) = "BI" THEN IULB_notes = "BI-Interface Prob"
		IF match_based_array(resolution_status_const,  item) = "BN" THEN IULB_notes = "BN-Already Known-No Savings"
		IF match_based_array(resolution_status_const,  item) = "BP" THEN IULB_notes = "BP-Wrong Person"
		IF match_based_array(resolution_status_const,  item) = "BU" THEN IULB_notes = "BU-Unable To Verify"
		IF match_based_array(resolution_status_const,  item) = "BO" THEN IULB_notes = "BO-Other"
		IF match_based_array(resolution_status_const,  item) = "NC" THEN IULB_notes = "NC-Non Cooperation"

		EMReadScreen panel_name, 4, 02, 52
	    IF panel_name = "IULB" THEN
	  		EMWriteScreen IULB_notes, 8, 6
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
			IF IULB_notes = "BE-No Change" THEN IULB_comment = "No change. " & other_notes
			IF IULB_notes = "BI-Interface Prob" THEN IULB_comment = "Interface Problem. " & other_notes
			IF IULB_notes = "BN-Already Known-No Savings" THEN IULB_comment = "Already known - No savings. " & other_notes
			IF IULB_notes = "BP-Wrong Person" THEN IULB_comment = "Client name and wage earner name are different. " & other_notes
			IF IULB_notes = "BU-Unable To Verify" THEN IULB_comment = "Unable To Verify. " & other_notes
			IF IULB_notes = "NC-Non Cooperation" THEN IULB_comment = "Non-coop, requested verf not in case file, " & other_notes

			IULB_comment = trim(IULB_comment)
			iulb_row = 8
			iulb_col = 6
			notes_array = split(IULB_comment, " ")
			For each word in notes_array
				EMWriteScreen word & " ", iulb_row, iulb_col
				If iulb_col + len(word) > 77 Then
					iulb_col = 6
					iulb_row = iulb_row + 1
					If iulb_row = 10 Then Exit For
				End If
				iulb_col = iulb_col + len(word) + 1
			Next
		   	TRANSMIT
			'------------------------------------------------------------------back on the IEVP menu, making sure that the match cleared
			EMReadScreen days_pending, 5, row, 72
	    	days_pending = trim(days_pending)

	    	IF IsNumeric(days_pending) = TRUE THEN
				match_based_array(date_cleared_const, item) = days_pending
			ELSE
				match_based_array(match_cleared_const, item) = TRUE 'match has now changed from match cleared False to True
				match_based_array(date_cleared_const, item) = date
                stats_counter = stats_counter + 1 'Increment for stats counter this will only count if true
			END IF
   		ELSE
			match_based_array(comments_const, item) = "Did not clear on IULB."
            'msgbox "7. Did not clear on IULB."
		END IF

        If match_based_array(match_cleared_const, item) = TRUE then
	 	    'Going to the MISC panel to add claim referral tracking information
	        '----------------------------------------------------------------------------------------------------CASE NOTE
		    CALL navigate_to_MAXIS_screen_review_PRIV("CASE", "NOTE", is_this_priv)
		    EMReadScreen county_code, 4, 21, 14  'Out of county cases from STAT
		    EMReadScreen case_invalid_error, 72, 24, 2 'if a person enters an invalid footer month for the case the script will attempt to  navigate'
		    case_invalid_error = trim(case_invalid_error)
		    IF priv_check = TRUE THEN  'PRIV cases
		        EMReadscreen priv_worker, 26, 24, 46
		        match_based_array(other_note_const, item) = "PRIV - Unable to case note "
		    ELSEIf county_code <> worker_county_code THEN
		      	match_based_array(other_note_const, item) = "OUT OF COUNTY CASE. Unable to case note."
		    ELSEIF instr(case_invalid_error, "IS INVALID") THEN  'CASE xxxxxxxx IS INVALID FOR PERIOD 12/99
		        match_based_array(other_note_const, item) = case_invalid_error & ". Unable to case note."
		    ELSE
		    	'-------------------------------------------------------------------for the case note
                IF match_type = "BEER" THEN match_type_letter = "B"
                IF match_type = "UBEN" THEN match_type_letter = "U"
                IF match_type = "UNVI" THEN match_type_letter = "U"
                IF match_type = "WAGE" THEN
                   IF select_quarter = 1 THEN IEVS_quarter = "1ST"
                   IF select_quarter = 2 THEN IEVS_quarter = "2ND"
                   IF select_quarter = 3 THEN IEVS_quarter = "3RD"
                   IF select_quarter = 4 THEN IEVS_quarter = "4TH"
                END IF
                IF match_type <> "UBEN" THEN IEVS_period = trim(replace(IEVS_period, "/", " to "))
                IF match_type = "UBEN" THEN IEVS_period = replace(IEVS_period, "-", "/")

		    	programs = ""
		        IF instr(match_based_array(program_const, item), "D") THEN programs = programs & "DWP, "
		        IF instr(match_based_array(program_const, item), "F") THEN programs = programs & "Food Support, "
		        IF instr(match_based_array(program_const, item), "H") THEN programs = programs & "Health Care, "
		        IF instr(match_based_array(program_const, item), "M") THEN programs = programs & "Medical Assistance, "
		        IF instr(match_based_array(program_const, item), "S") THEN programs = programs & "MFIP, "
		        'trims excess spaces of programs
		        programs = trim(programs)
		        'takes the last comma off of programs when autofilled into dialog
		        IF right(programs, 1) = "," THEN programs = left(programs, len(programs) - 1)

    	        PF9
                'Case note header options based on the match type
    	        IF match_type = "WAGE" THEN
                    CALL write_variable_in_case_note("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE MATCH"  & " (" & first_name & ") CLEARED " & match_based_array(resolution_status_const,  item) & "-----")
                Elseif match_type = "BNDX" THEN
		    		CALL write_variable_in_case_note("-----" & IEVS_period & " NON-WAGE MATCH(" & match_type & ")" & " (" & first_name & ") CLEARED " & match_based_array(resolution_status_const,  item) & "-----")
                ELSE
                	CALL write_variable_in_case_note("-----" & IEVS_year & " NON-WAGE MATCH(" & match_type_letter & ")" & " (" & first_name & ") CLEARED " & match_based_array(resolution_status_const,  item) & "-----")
    	    	END IF
    	        CALL write_bullet_and_variable_in_case_note("Period", match_based_array(period_const, item))
    	        CALL write_bullet_and_variable_in_case_note("Programs on Match", programs)
				CALL write_bullet_and_variable_in_case_note("Active Programs", list_active_programs)
				CALL write_bullet_and_variable_in_case_note("Pending Programs", list_pending_programs)
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
				CALL write_variable_in_case_note(worker_signature)
    	        PF3 'to save casenote'
    	    	match_based_array(comments_const, item) = "Match Cleared and Case Noted."
                'msgbox "8. Match Cleared and Case Noted."
		    END IF
        END IF
	END IF
NEXT

'Excel headers and formatting the columns
objExcel.Cells(1, 1).Value     = "DATE POSTED" 		'A Date Posted to Maxis'
objExcel.Cells(1, 2).Value     = "BASKET" 			'B Worker #
objExcel.Cells(1, 3).Value     = "DOB" 				'C DOB
objExcel.Cells(1, 4).Value     = "RELATIONSHIP" 	'D Relationship
objExcel.Cells(1, 5).Value     = "CASE NUMBER" 		'E Maxis case number
objExcel.Cells(1, 6).Value     = "EARNER NAME" 		'F Earner Name
objExcel.Cells(1, 7).Value     = "CLIENT NAME" 		'G Case Name
objExcel.Cells(1, 8).Value     = "SSN" 				'H SSN
objExcel.Cells(1, 9).Value     = "PROG"			    'I Program
objExcel.Cells(1, 10).Value    = "AMOUNT"			'J Amount
objExcel.Cells(1, 11).Value    = "SOURCE OF INCOME" 'K Employer
objExcel.Cells(1, 12).Value    = "NOTICE SENT"		'L Date Notice Sent
objExcel.Cells(1, 13).Value    = "RESOLUTION"		'M How cleared
objExcel.Cells(1, 14).Value    = "DATE CLEARED"		'N Date cleared
objExcel.Cells(1, 15).Value    = "CLAIM #"			'P Claim Entered
objExcel.Cells(1, 16).Value    = "ASSIGNED TO"		'O Worker who cleared
objExcel.Cells(1, 17).Value    = "MATCH TYPE"    	'Q Match type number
objExcel.Cells(1, 18).Value    = "PERIOD"	        'R Match periods
objExcel.Cells(1, 19).Value    = "DATE ATR RCVD"	'S Date ATR on file
objExcel.Cells(1, 20).Value    = "DATE EVF SIGNED"	'T Date EVF Received
objExcel.Cells(1, 21).Value    = "OTHER NOTES"		'U Other Notes
objExcel.Cells(1, 22).Value    = "COMMENTS"		    'V Comments

MsgBox "Writing to excel- please dont touch the keyboard until list is fully updated. You'll receive a closing success message." 'this is working as a ready wait'

For item = 0 to UBound(match_based_array, 2)
 	excel_row = match_based_array(excel_row_const, item)
 	objExcel.Cells(excel_row, excel_col_comments).Value 	= match_based_array(comments_const, item)
	objExcel.Cells(excel_row, excel_date_notice_sent).Value	= match_based_array(notice_sent_date_const, item)
	objExcel.Cells(excel_row, excel_col_date_cleared).Value = match_based_array(date_cleared_const, item)
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

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------03/11/2022
'--Tab orders reviewed & confirmed----------------------------------------------03/11/2022
'--Mandatory fields all present & Reviewed--------------------------------------03/11/2022
'--All variables in dialog match mandatory fields-------------------------------03/11/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------03/11/2022
'--CASE:NOTE Header doesn't look funky------------------------------------------03/11/2022
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------03/11/2022
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used-11/15/2022
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------03/11/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------03/11/2022
'--PRIV Case handling reviewed -------------------------------------------------03/11/2022
'--Out-of-County handling reviewed----------------------------------------------03/11/2022
'--script_end_procedures (w/ or w/o error messaging)----------------------------03/11/2022
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------03/11/2022
'--Incrementors reviewed (if necessary)-----------------------------------------03/11/2022
'--Denomination reviewed -------------------------------------------------------03/11/2022
'--Script name reviewed---------------------------------------------------------03/11/2022 purposed rename BULK-DEU-MATCH CLEARED
'--BULK - remove 1 incrementor at end of script reviewed------------------------03/11/2022

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub taks are complete-----------------------------------------03/11/2022
'--Comment Code-----------------------------------------------------------------03/11/2022
'--Update Changelog for release/update------------------------------------------03/11/2022
'--Remove testing message boxes-------------------------------------------------03/11/2022
'--Remove testing code/unnecessary code-----------------------------------------03/11/2022
'--Review/update SharePoint instructions----------------------------------------03/11/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------03/11/2022
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------N/A
'--Complete misc. documentation (if applicable)---------------------------------03/11/2022
'--Update project team/issue contact (if applicable)----------------------------03/11/2022
