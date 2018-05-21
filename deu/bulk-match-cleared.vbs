'GATHERING STATS===========================================================================================
name_of_script = "BULK-MATCH-CLEARED.vbs"
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
			BeginDialog IEVS_match_dialog, 0, 0, 266, 140, "BULK IEVS Match"
  				DropListBox 200, 50, 50, 15, "Select one..."+chr(9)+"BEER"+chr(9)+"WAGE", IEVS_type
           		ButtonGroup ButtonPressed
            	PushButton 200, 70, 50, 15, "Browse...", select_a_file_button
            	OkButton 145, 115, 50, 15
            	CancelButton 200, 115, 50, 15
           		EditBox 15, 70, 180, 15, IEVS_match_path
           		Text 20, 30, 235, 20, "This script should be used when IEVS matches have been researched and ready to be cleared. "
           		Text 20, 90, 230, 15, "Select the Excel file that contains the case inforamtion by selecting the 'Browse' button, and finding the file."
           		Text 55, 55, 135, 10, "Select the type of IEVS match to process:"
           		GroupBox 10, 10, 250, 100, "Using the IEVS match script"
         	EndDialog
			err_msg = ""
			Dialog IEVS_match_dialog
			cancel_confirmation
			If ButtonPressed = select_a_file_button then
				If IEVS_match_path <> "" then 'This is handling for if the BROWSE button is pushed more than once'
					objExcel.Quit 'Closing the Excel file that was opened on the first push'
					objExcel = "" 	'Blanks out the previous file path'
				End If
				call file_selection_system_dialog(IEVS_match_path, ".xlsx") 'allows the user to select the file'
			End If
			If IEVS_type = "Select one..." then err_msg = err_msg & vbNewLine & "* Select type of match you are processing."
			If IEVS_match_path = "" then err_msg = err_msg & vbNewLine & "* Use the Browse Button to select the file that has your client data"
			If err_msg <> "" Then MsgBox err_msg
		Loop until err_msg = ""
		If objExcel = "" Then call excel_open(IEVS_match_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
		If err_msg <> "" Then MsgBox err_msg
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

excel_row = 2			'establishing row to start
DO
	MAXIS_case_number 	= objExcel.cells(excel_row, 1).value	'establishes MAXIS case number
	Client_SSN 			= objExcel.cells(excel_row, 3).value	'establishes client SSN
	Employer_name		= objExcel.cells(excel_row, 6).value	'establishes employer name
	Cleared_status	    = objExcel.cells(excel_row, 8).value	'establishes cleared status for the match
	'cleaned up
	MAXIS_case_number 	= trim(MAXIS_case_number) 'remove extra spaces'
	Client_SSN 			= trim(Client_SSN)
	Employer_name 	   	= trim(Employer_name)
	Cleared_status 	  	= trim(Cleared_status)

    If MAXIS_case_number = "" then exit do 'goes to actions outside of do loop'
	back_to_self
	'----------------------------------------------------------------------------------------------------DAIL
	Call navigate_to_MAXIS_screen("DAIL", "DAIL")
	'Making sure that the user is on an acceptable DAIL message
	EMReadScreen case_number, 8, 5, 73
	case_number = trim(case_number)
	IF case_number <> MAXIS_case_number then
		EMreadscreen case_number, 8, 7, 73   'DAILS often read down two check to see if matching'
		 If case_number <> MAXIS_case_number then
			objExcel.cells(excel_row, 9).value = "A pending IEVS match could not be found on DAIL/DAIL."
			match_found = False
		End if
	Else
	    row = 6    'establishing 1st row to search
	    Do
		    EMReadScreen IEVS_message, 4, row, 6
		    'msgbox IEVS_message & vbcr & IEVS_type
		    If IEVS_message <> IEVS_type then
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
	End if 			
		
	If match_found = False then 
		case_note_actions = False 'no case note'
		objExcel.cells(excel_row, 9).value = "A IEVS match wasn't found on DAIL/DAIL or SSN did not match."
	End if 
	'----------------------------------------------------------------------------------------------------IEVS
	If match_found = True then
	    'Navigating deeper into the match interface
	    CALL write_value_and_transmit("I", row, 3)   'navigates to INFC
	    CALL write_value_and_transmit("IEVP", 20, 71)   'navigates to IEVP
		EMReadScreen error_msg, 7, 24, 2
		If error_msg = "NO IEVS" then 'checking for error msg'
			objExcel.cells(excel_row, 9).value = "No IEVS matches found for SSN " & Client_SSN & "/Could not access IEVP."
			case_note_actions = False
		Else
			row = 7
		    'Ensuring that match has not already been resolved.
		    Do
				EMReadScreen days_pending, 5, row, 72
		    	days_pending = trim(days_pending)
		    	If IsNumeric(days_pending) = false then
					objExcel.cells(excel_row, 9).value = "No pending IEVS match found. Please review IEVP."
					case_note_actions = False
					exit do
		    	ELSE
	            	'Entering the IEVS match & reading the difference notice to ensure this has been sent
                	EMReadScreen IEVS_period, 11, row, 47
		    		EMReadScreen start_month, 2, row, 47
		    		EMReadScreen end_month, 2, row, 53
					If trim(start_month) = "" or trim(end_month) = "" then
						case_note_actions = False
		    		else
						month_difference = abs(end_month) - abs(start_month)
					    If (IEVS_type = "WAGE" and month_difference = 2) then 'ensuring if it is a wage the match is a quater'
					    	case_note_actions = true
					    	exit do
					    Elseif (IEVS_type = "BEER" and month_difference = 11) then  'ensuring that if it a beer that the match is a year'
					    	case_note_actions = True
					    	exit do
					    End if
					End if
					row = row + 1
				END IF
			Loop until row = 17

			If case_note_actions <> True then
				If IEVS_type = "WAGE" then
			    	objExcel.cells(excel_row, 9).value = "This WAGE match is not for a quarter. Please process manually."
			    Elseif IEVS_type = "BEER" then
					objExcel.cells(excel_row, 9).value = "This BEER match is not for a year. Please process manually."
				END if
				case_note_actions = False
			Else
		        CALL write_value_and_transmit("U", row, 3)   'navigates to IULA
                'Reading the IEVS period to be autofilled into the cleared_match_dialog

				'Reading potential errors for out-of-county cases
				EMReadScreen OutOfCounty_error, 12, 24, 2
				IF OutOfCounty_error = "MATCH IS NOT"  then
					objExcel.cells(excel_row, 9).value = "Out-of-county case. Cannot update."
					case_note_actions = False
				else
                    IF IEVS_type = "WAGE" then
				    	EMReadScreen quarter, 1, 8, 14
                    	EMReadScreen IEVS_year, 4, 8, 22
				    Elseif IEVS_type = "BEER" then
				    	EMReadScreen IEVS_year, 2, 8, 15
				    	IEVS_year = "20" & IEVS_year
				    End if

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

                    EMReadScreen diff_date, 10, 14, 68
	                diff_date = trim(diff_date)
		            If diff_date <> "" then
			        	diff_date = replace(diff_date, " ", "/") 'replace spaces with format to date'
                        objExcel.cells(excel_row, 7).value = diff_date
                    END IF
					
					EMReadScreen Active_Programs, 13, 6, 68
					Active_Programs =trim(Active_Programs)
					objExcel.cells(excel_row, 4).value = Active_Programs	
					
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
						EMReadScreen cleared_field, 2, row + 1, col + 1
						If cleared_field <> "__" then
							objExcel.cells(excel_row, 9).value = "Unable to update cleared status on IULA."
							case_note_actions = False
						Else 
							EMWriteScreen Cleared_status, row + 1, col + 1
						End if 
					Next 
					
                    CALL write_value_and_transmit("10", 12, 46)   'navigates to IULB

				    'resolved notes depending on the Cleared_status
				    If Cleared_status = "BC" then CALL write_value_and_transmit("Case closed.", 8, 6)   'BC
                    If Cleared_status = "BE" then CALL write_value_and_transmit("No change.", 8, 6)   'BE
				    If Cleared_status = "BN" then CALL write_value_and_transmit("Already known - No savings.", 8, 6)   'BN
				    If Cleared_status = "CC" then CALL write_value_and_transmit("Claim entered.", 8, 6)   'CC
					objExcel.cells(excel_row, 9).value = "IEVS match cleared"
                    case_note_actions = True
				End if
			End if
		End if
	End if

    If case_note_actions = True then		'Formatting for the case note
	    If IEVS_type = "WAGE" then
	    	'Updated IEVS_period to write into case note
	    	If quarter = 1 then IEVS_quarter = "1ST"
	    	If quarter = 2 then IEVS_quarter = "2ND"
	    	If quarter = 3 then IEVS_quarter = "3RD"
	    	If quarter = 4 then IEVS_quarter = "4TH"
	    End if

	    'adding specific wording for case note header for each cleared status
	    If Cleared_status = "BC" then cleared_header_info = " (" & first_name & ") CLEARED BC-CASE CLOSED"
	    If Cleared_status = "BE" then cleared_header_info = " (" & first_name & ") CLEARED BE-NO CHANGE"
	    If Cleared_status = "BN" then cleared_header_info = " (" & first_name & ") CLEARED BN-KNOWN"
	    If Cleared_status = "CC" then cleared_header_info = " (" & first_name & ") CLEARED CC-CLAIM ENTERED"

		'Case noting the actions taken
        start_a_blank_CASE_NOTE
        If IEVS_type = "WAGE" then Call write_variable_in_CASE_NOTE("-----" & IEVS_quarter & " QTR " & IEVS_year & " WAGE INCOME" & cleared_header_info & "-----")
		If IEVS_type = "BEER" then Call write_variable_in_CASE_NOTE("-----" & IEVS_year & " NON WAGE INCOME(B)" & cleared_header_info & "-----")
		Call write_bullet_and_variable_in_CASE_NOTE("Period", IEVS_period)
		Call write_bullet_and_variable_in_CASE_NOTE("Programs open", programs)
        Call write_bullet_and_variable_in_CASE_NOTE("Employer name", Employer_name)
        call write_variable_in_CASE_NOTE("------ ----- -----")
        If Cleared_status = "BN" or Cleared_status = "BE" then Call write_variable_in_CASE_NOTE("CLIENT REPORTED EARNINGS. INCOME IS IN STAT/JOBS AND BUDGETED.")
        If Cleared_status <> "CC" then Call write_variable_in_CASE_NOTE("NO OVERPAYMENTS OR SAVINGS RELATED TO THIS MATCH.")
        call write_variable_in_CASE_NOTE("------ ----- ----- ----- -----")
        Call write_variable_in_CASE_NOTE("DEBT ESTABLISHMENT UNIT 612-348-4290 PROMPTS 1-1-1")
    End if
	
	excel_row = excel_row + 1
	MAXIS_case_number = ""
	Client_SSN = ""
	STATS_counter = STATS_counter + 1
LOOP UNTIL objExcel.Cells(excel_row, 1).value = ""	'looping until the list of cases to check for recert is complete

STATS_counter = STATS_counter - 1		'removes 1 to correct the count
script_end_procedure("Success! The IEVS match cases have now been updated. Please review the NOTES section to review the cases/follow up work to be completed.")