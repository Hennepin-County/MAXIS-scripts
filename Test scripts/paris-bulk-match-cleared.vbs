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
call changelog_update("03/20/2020", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display

'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to MAXIS
EMConnect ""

Do
	Do
		'The dialog is defined in the loop as it can change as buttons are pressed
		'-------------------------------------------------------------------------------------------------DIALOG
		Dialog1 = "" 'Blanking out previous dialog detail
		BeginDialog Dialog1, 0, 0, 271, 155, "BULK-Paris Match Cleared"
		  EditBox 65, 20, 195, 15, other_notes
		  ButtonGroup ButtonPressed
		    PushButton 10, 85, 50, 15, "Browse:", select_a_file_button
		  EditBox 65, 85, 195, 15, IEVS_match_path
		  ButtonGroup ButtonPressed
		    OkButton 170, 135, 45, 15
		    CancelButton 220, 135, 45, 15
		  Text 10, 25, 45, 10, "Other Notes:"
		  GroupBox 5, 50, 260, 80, "Using the script:"
		  Text 10, 60, 250, 15, "Select the Excel file that contains the case information by selecting the 'Browse' button and locating the file."
		  Text 10, 105, 245, 15, "This script should be used when matches have been researched and ready to be cleared. "
		  GroupBox 5, 5, 260, 45, "Complete prior to browsing the script:"
		  Text 65, 40, 145, 10, "This will be in the case note of each case"
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
		If select_match_type = "Select One:" then err_msg = err_msg & vbNewLine & "* Select type of match you are processing."
		If IEVS_match_path = "" then err_msg = err_msg & vbNewLine & "* Use the Browse Button to select the file that has your client data"
		If err_msg <> "" Then MsgBox err_msg
	Loop until err_msg = ""
	If objExcel = "" Then call excel_open(IEVS_match_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
	If err_msg <> "" Then MsgBox err_msg
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

excel_row = 2			'establishing row to start
DO
	excel_state_match					= objExcel.cells(excel_row, 6).value	'establishes MAXIS case number
	MAXIS_case_number 					= objExcel.cells(excel_row, 7).value	'establishes MAXIS case number
    Client_SSN 							= objExcel.cells(excel_row, 8).value	'establishes client SSN
	first_name 							= objExcel.cells(excel_row, 10).value	'establishes client SSN
    excel_date_notice_sent				= objExcel.cells(excel_row, 11).value	'establishes if notice has already been sent
    worker_entered_cleared_status	    = objExcel.cells(excel_row, 13).value	'establishes cleared status for the match
    excel_date_contact_with_other_state = objExcel.cells(excel_row, 12).value	'establishes cleared status for the match
    excel_date_cleared					= objExcel.cells(excel_row, 14).value
	'TODO other_notes so that workers can have a single note on a case
    'cleaned up
	excel_state_match					= trim(excel_state_match) 'remove extra spaces'
    MAXIS_case_number 					= trim(MAXIS_case_number)
    Client_SSN 							= trim(Client_SSN)
    Client_SSN 							= replace(Client_SSN, "-", "")
	first_name 							= trim(first_name)
    excel_date_notice_sent 				= trim(excel_date_notice_sent)
    worker_entered_cleared_status 	  	= trim(worker_entered_cleared_status)

    IF worker_entered_cleared_status <> "PR" or worker_entered_cleared_status <> "HM" or worker_entered_cleared_status <> "RV" or worker_entered_cleared_status <> "FR" or worker_entered_cleared_status <> "CC" THEN err_msg = err_msg & vbNewLine & "Please only use applicable codes when clearing a match."
    'UR Unresolved, System Entered Only
    'PR Person Removed From Household
    'HM Household Moved Out Of State
    'RV Residency Verified, Person in MN
    'FR Failed Residency Verification Request
    'PC Person Closed, Not PARIS Interstate
    'CC Case Closed, Not PARIS Interstate
    If MAXIS_case_number = "" THEN exit do 'goes to actions outside of do loop'
    back_to_self
	'----------------------------------------------------------------------------------------------------DAIL
	Call navigate_to_MAXIS_screen("DAIL", "DAIL")
	'Making sure that the user is on an acceptable DAIL message
	EMReadScreen case_number, 8, 5, 73
	case_number = trim(case_number)
	IF case_number <> MAXIS_case_number then
		EMreadscreen case_number, 8, 7, 72   'DAILS often read down two check to see if matching'
		If case_number <> MAXIS_case_number then
			objExcel.cells(excel_row, 17).value = "A pending match could not be found on DAIL/DAIL."
			match_found = FALSE
			case_note_actions = FALSE
		End if
	Else
	    row = 6    'establishing 1st row to search
	    Do
		    EMReadScreen IEVS_message, 4, row, 6
		    'msgbox IEVS_message & vbcr & select_match_type
		    If trim(IEVS_message) <> "PARI" then
				match_found = FALSE
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
				'msgbox client_social & first_name & Client_SSN

		    	If client_social <> Client_SSN then
		    		match_found = FALSE
		    		row = row + 1
					'msgbox "2." & MAXIS_case_number & vbcr & "row: " & row & vbcr & "match found: " & match_found
		    	Else
		    		match_found = true
					'msgbox "3." & MAXIS_case_number & vbcr & "row: " & row & vbcr & "match found: " & match_found
		    		exit do
		    	End if
		    End if
		Loop until match_found = true or row = 19
		If match_found = FALSE then
			case_note_actions = FALSE 'no case note'
			objExcel.cells(excel_row, 17).value = "A match wasn't found on DAIL/DAIL or SSN did not match."
		End if
	End if

	'----------------------------------------------------------------------------------------------------IEVS
	If match_found = True then
	    'Navigating deeper into the match interface
	    CALL write_value_and_transmit("I", row, 3)   'navigates to INFC
	    CALL write_value_and_transmit("INTM", 20, 71)   'navigates to IEVP
     	EMReadScreen error_msg, 7, 24, 2
		If error_msg = "NO IEVS" then 'checking for error msg'
			objExcel.cells(excel_row, 17).value = "No matches found for SSN " & Client_SSN & "/Could not access INTM."
			case_note_actions = FALSE
		Else
			row = 8
		    'Ensuring that match has not already been resolved.
		    Do
				EMReadScreen INTM_match_status, 2, row, 73 'DO loop to check status of case before we go into insm'
		    	INTM_match_status = trim(INTM_match_status)
				'msgbox INTM_match_status
				EMReadScreen INTM_period, 5, row, 59
		    	If INTM_match_status = "" THEN
					objExcel.cells(excel_row, 17).value = "No pending match found. Please review INTM."
					case_note_actions = FALSE
					exit do
		    	ELSE
					If INTM_match_status = "UR" THEN
						case_note_actions = TRUE
						exit do
					End if
					row = row + 1
				END IF
			Loop until row = 17

			'---------------------------------------------------------------------Reading potential errors for out-of-county cases

			CALL write_value_and_transmit("X", row, 3)   'navigates to IULA
			EMReadScreen OutOfCounty_error, 12, 24, 2
			IF OutOfCounty_error = "MATCH IS NOT" then
				objExcel.cells(excel_row, 17).value = "Out-of-county case. Cannot update."
				case_note_actions = FALSE
			ELSE
			    'Ensuring that the client has not already had a difference notice sent
			    EMReadScreen notice_sent, 1, 8, 73
			    EMReadScreen notice_sent_date, 8, 9, 73
			    notice_sent_date = trim(notice_sent_date)
			    If trim(notice_sent_date) <> "" then notice_sent_date= replace(notice_sent_date, " ", "/")
			END IF
			If notice_sent = "Y" THEN
				objExcel.cells(excel_row, 17).value = "Notice Sent"
				objExcel.cells(excel_row, 11).value = notice_sent_date
				case_note_actions = FALSE
				EXIT DO
			END IF
			'--------------------------------------------------------------------Client name
			EmReadScreen panel_name, 4, 02, 55
			IF panel_name <> "INSM" THEN
				objExcel.cells(excel_row, 17).value = "Script did not find INSM"
				EXIT DO
			ELSE'----------------------------------------------------------------------Minnesota active programs
			    EMReadScreen MN_MN_active_programs, 15, 6, 59
			    MN_active_programs = Trim(MN_active_programs)
			    MN_active_programs = replace(MN_active_programs, " ", ", ")
			    programs = ""
			    IF instr(MN_active_programs, "CASH") THEN programs = programs & "CASH, "
			    IF instr(MN_active_programs, "FS") THEN programs = programs & "FOOD SUPPORT, "
			    IF instr(MN_active_programs, "HC") THEN programs = programs & "HEALTH CARE, "
			    IF instr(MN_active_programs, "MCRE") THEN programs = programs & "MinnesotaCare, "

			    'trims excess spaces of programs
			    programs = trim(programs)
			    'takes the last comma off of programs when autofilled into dialog
			    IF right(programs, 1) = "," THEN programs = left(programs, len(programs) - 1)
			    'Month of the PARIS match
			    EMReadScreen Match_Month, 5, 6, 27
			    Match_month = replace(Match_Month, " ", "/")

			    '--------------------------------------------------------------------PARIS match state & active programs-this will handle more than one state
			row = 13
			DO
				'-------------------------------------------------------Reading for each state active programs
				EMReadScreen match_state, 2, row, 3
				IF trim(match_state) = "" THEN
					EXIT DO
				ELSE  '-------------------------------------------------------------------Case number for match state (if exists)
					IF match_state = excel_state_match THEN
					   	EMReadScreen Match_State_Case_Number, 13, row, 9
					    Match_State_Case_Number = trim(Match_State_Case_Number)
					    IF Match_State_Case_Number = "" THEN Match_State_Case_Number = "N/A"
						'-------------------------------------------------------------------PARIS match contact information
						EMReadScreen phone_number, 23, row, 22
						phone_number = TRIM(phone_number)
						If phone_number = "Phone: (     )" then
							phone_number = ""
						Else
							EMReadScreen phone_number_ext, 8, row, 51
							phone_number_ext = trim(phone_number_ext)
							If phone_number_ext <> "" then phone_number = phone_number & " Ext: " & phone_number_ext
						End if

						'-------------------------------------------------------------------reading and cleaning up the fax number if it exists
						EMReadScreen fax_check, 8, row + 1, 31
						fax_check = trim(fax_check)
						If fax_check <> "" then
							EMReadScreen fax_number, 21, row + 1, 24
							fax_number = TRIM(fax_number)
						End if
						If fax_number = "Fax: (     )" then fax_number = ""
						match_contact_info = phone_number & " " & fax_number
						'-------------------------------------------------------------------trims excess spaces of match_programs
						match_programs = "" 'sometimes blanking over information will clear the value of the variable'
						'match_row = row           'establishing match row the same as the current state row. Needs another variables since we are only incrementing the match row in the loop. Row needs to stay the same for larger loop/next state.
						DO
							EMReadScreen match_state_active_programs, 22, row, 60
							match_state_active_programs = TRIM(match_state_active_programs)
							IF match_state_active_programs = "" THEN EXIT DO
							IF match_state_active_programs = "FOOD SUPPORT" THEN match_programs = match_programs & "FS, "
							IF match_state_active_programs = "HEALTH CARE" THEN match_programs = match_programs &  "HC, "
							IF match_state_active_programs = "STATE SSI" THEN match_programs = match_programs & "SSI, "
							IF match_state_active_programs = "NONE IDICATED" THEN match_programs = match_programs &  "NONE INDICATED"
							IF match_state_active_programs = "CASH" THEN match_programs = match_programs &  "CASH"
							IF match_state_active_programs = "CHILD CARE" THEN match_programs = match_programs &  "CCA"
							IF match_state_active_programs = "STATE WORKERS COMP" THEN match_programs = match_programs &  "WORKERS COMP"
							row = row + 1
						LOOP

						'trims excess spaces of programs
						match_programs = trim(match_programs)
						'takes the last comma off of programs when autofilled into dialog
						IF right(match_programs, 1) = "," THEN match_programs = left(match_programs, len(match_programs) - 1)
						'MsgBox match_programs
					END IF
 					row = row + 3
					EMReadScreen add_match_state, 2, row, 3
					add_match_state = trim(add_match_state)
					MsgBox "what the h?" & row & add_match_state & second_match
					IF trim(add_match_state) = "" THEN
						EXIT DO
						second_match = FALSE
						MsgBox add_match_state & second_match
					ELSE
						IF add_match_state <> "" THEN
							MsgBox add_match_state
					   		EMReadScreen add_state_Case_Number, 13, row, 9
					    	add_state_Case_Number = trim(add_state_Case_Number)
					    	IF add_state_Case_Number = "" THEN add_state_Case_Number = "N/A"
							'-------------------------------------------------------------------PARIS match contact information
							EMReadScreen add_state_phone_number, 23, row, 22
							add_state_phone_number = TRIM(add_state_phone_number)
							If add_state_phone_number = "Phone: (     )" then
								add_state_phone_number = ""
							Else
								EMReadScreen add_state_phone_number_ext, 8, row, 51
								add_state_phone_number_ext = trim(add_state_phone_number_ext)
								If add_state_phone_number_ext <> "" then add_state_phone_number = add_state_phone_number & " Ext: " & add_state_phone_number_ext
							End if

							'-------------------------------------------------------------------reading and cleaning up the fax number if it exists
							EMReadScreen add_state_fax_check, 8, row + 1, 31
							add_state_fax_check = trim(add_state_fax_check)
							If add_state_fax_check <> "" then
								EMReadScreen add_state_fax_number, 21, row + 1, 24
								add_state_fax_number = TRIM(add_state_fax_number)
							End if
							If add_state_fax_number = "Fax: (     )" then add_state_fax_number = ""
							match_contact_info = add_state_phone_number & " " & add_state_fax_number

						    '-------------------------------------------------------------------trims excess spaces of match_programs
						    DO
						    	EMReadScreen add_match_state_active_programs, 22, row, 60
						    	add_match_state_active_programs = TRIM(add_match_state_active_programs)
						    	IF add_match_state_active_programs = "" THEN EXIT DO
						    	IF add_match_state_active_programs = "FOOD SUPPORT" THEN match_programs = add_match_programs & "FS, "
						    	IF add_match_state_active_programs = "HEALTH CARE" THEN add_match_programs = add_match_programs &  "HC, "
						    	IF add_match_state_active_programs = "STATE SSI" THEN add_match_programs = add_match_programs & "SSI, "
						    	IF add_match_state_active_programs = "NONE IDICATED" THEN add_match_programs = add_match_programs &  "NONE INDICATED"
						    	IF add_match_state_active_programs = "CASH" THEN add_match_programs = add_match_programs &  "CASH"
						    	IF add_match_state_active_programs = "CHILD CARE" THEN add_match_programs = add_match_programs &  "CCA"
						    	IF add_match_state_active_programs = "STATE WORKERS COMP" THEN add_match_programs = add_match_programs &  "WORKERS COMP"
						    	row = row + 1
						    	MsgBox add_match_programs
						    LOOP
						    'trims excess spaces of programs
						    add_match_programs = trim(add_match_programs)
						    'takes the last comma off of programs when autofilled into dialog
						    IF right(add_match_programs, 1) = "," THEN add_match_programs = left(add_match_programs, len(add_match_programs) - 1)
						END IF
					END IF
					row = row + 1
					IF row = 19 THEN
						PF8
						EMReadScreen last_page_check, 21, 24, 2
						last_page_check = trim(last_page_check)
						IF last_page_check = ""  THEN row = 13
					END IF
				END IF
			LOOP UNTIL last_page_check = "THIS IS THE LAST PAGE"

			IF notice_sent = "N" and worker_entered_cleared_status = "" and case_note_actions = TRUE THEN 'sending the notice
			    PF9	'edit mode'
				MsgBox "we want to send a notice"
			    EMReadScreen edit_error, 2, 24, 2
			    edit_error = trim (edit_error)
			    IF edit_error <> "" THEN script_end_procedure ("Unable to send difference notice please review case")
			    EMwritescreen "Y", 8, 73 'send Notice
			    TRANSMIT
			    '--------------------------------------------------------------------The case note & case note related code
			    Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days
				pending_verifs = "Interstate Match Notice, Shelter Verification, Proof of Residency"

			    '-----------------------------------------------------------------------The case note
			    start_a_blank_CASE_NOTE
			    CALL write_variable_in_CASE_NOTE ("-----" & Match_month & " PARIS MATCH " & "(" & first_name &  ") NOTICE SENT-----")
			    CALL write_bullet_and_variable_in_CASE_NOTE("Client Name", Client_Name)
			    CALL write_bullet_and_variable_in_CASE_NOTE("MN Active Programs", MN_active_programs)
				CALL write_variable_in_CASE_NOTE("-----Match State: " & excel_state_match & "-----")
			    CALL write_bullet_and_variable_in_CASE_NOTE("Match State Active Programs", match_programs)
			    CALL write_bullet_and_variable_in_CASE_NOTE("Match State Contact Info", match_state_contact_info )
				IF add_match_state <> "" THEN
					CALL write_variable_in_CASE_NOTE("-----Match State: " & add_match_state & "-----")
			    	CALL write_bullet_and_variable_in_CASE_NOTE("Match State Active Programs", add_match_programs)
			    	CALL write_bullet_and_variable_in_CASE_NOTE("Match State Contact Info:", add_match_contact_info )
				END IF
			    CALL write_variable_in_CASE_NOTE ("-----")
			    CALL write_bullet_and_variable_in_CASE_NOTE("Contacted other state", excel_date_contact_with_other_state)
			    CALL write_bullet_and_variable_in_CASE_NOTE("Verification Requested", pending_verifs)
			    CALL write_bullet_and_variable_in_CASE_NOTE("Verification Due", Due_date)
			    CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
			    CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- -----")
			    CALL write_variable_in_CASE_NOTE ("DEBT ESTABLISHMENT UNIT 612-348-4290 EXT 1-1-1")
				objExcel.cells(excel_row, 17).value = "Notice Sent"
				objExcel.cells(excel_row, 11).value = DATE

			ELSEif worker_entered_cleared_status <> "" THEN
				msgbox "We want to clear"

			    Due_date = dateadd("d", 10, date)	'defaults the due date for all verifications at 10 days
			    'requested for HEADER of casenote'

				'IF worker_entered_cleared_status = "PR" THEN rez_status = "PR-Person Removed From Household"
				'IF worker_entered_cleared_status = "HM" THEN rez_status = "HM-Household Moved Out Of State"
				'IF worker_entered_cleared_status = "RV" THEN rez_status = "RV-Residency Verified, Person in MN"
				'IF worker_entered_cleared_status = "FR" THEN rez_status = "FR-Failed Residency Verification Request"
				'IF worker_entered_cleared_status = "PC" THEN rez_status = "PC-Person Closed, Not PARIS Interstate"
				'--------------------------------'still need to be on PARIS Interstate Match Display (INSM)'
				PF9
				EMwritescreen worker_entered_cleared_status, 9, 27
				IF fraud_referral_excel = "YES" THEN
					EMwritescreen "Y", 10, 27
					ELSE
					TRANSMIT
				END IF
				MsgBox "did we clear?"
				PF3
				PF3
			    '----------------------------------------------------------------the case match note
			    start_a_blank_CASE_NOTE
			    CALL write_variable_in_CASE_NOTE ("-----" & Match_month & " PARIS MATCH " & "(" & first_name &  ") CLEARED " & worker_entered_cleared_status & "-----")
			    CALL write_bullet_and_variable_in_CASE_NOTE("Client Name", Client_Name)
			    CALL write_bullet_and_variable_in_CASE_NOTE("MN Active Programs", MN_active_programs)
				Call write_bullet_and_variable_in_case_note("Discovery date", contact_other_state)
				Call write_bullet_and_variable_in_case_note("Period", INTM_period)
				CALL write_variable_in_CASE_NOTE("-----Match State: " & excel_state_match & "-----")
			    CALL write_bullet_and_variable_in_CASE_NOTE("Match State Active Programs", match_programs)
			    CALL write_bullet_and_variable_in_CASE_NOTE("Match State Contact Info", match_state_contact_info )
				IF add_match_state <> "" THEN
					CALL write_variable_in_CASE_NOTE("-----Match State: " & add_match_state & "-----")
			    	CALL write_bullet_and_variable_in_CASE_NOTE("Match State Active Programs", add_match_programs)
			    	CALL write_bullet_and_variable_in_CASE_NOTE("Match State Contact Info", add_match_contact_info )
				END IF
			    CALL write_variable_in_CASE_NOTE ("-----")
			    'CALL write_bullet_and_variable_in_CASE_NOTE("Client accessing benefits in other state", bene_other_state)
			    CALL write_bullet_and_variable_in_CASE_NOTE("Contacted other state", excel_date_contact_with_other_state)
			    'CALL write_bullet_and_variable_in_CASE_NOTE("Verification used to clear", pending_verifs)
			    'CALL write_bullet_and_variable_in_CASE_NOTE("Resolution Status", resolution_status)
				'IF worker_entered_cleared_status = "FR" THEN
					'CALL write_variable_in_CASE_NOTE("Client has failed to cooperate with Paris Match - has not provided requested verifications showing they are living in MN. Client will need to provide this before the case is reopened ")
				'END IF
				'CALL write_bullet_and_variable_in_case_note("Fraud referral made", fraud_referral)
			    CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
			    CALL write_variable_in_CASE_NOTE("----- ----- ----- ----- ----- ----- -----")
			    CALL write_variable_in_CASE_NOTE ("DEBT ESTABLISHMENT UNIT 612-348-4290 EXT 1-1-1")
				objExcel.cells(excel_row, 11).value = notice_sent_date
				objExcel.cells(excel_row, 17).value = worker_entered_cleared_status
			END IF
		END IF
		END IF
	END IF
	excel_row = excel_row + 1

	excel_state_match					= ""
	MAXIS_case_number 					= ""
	Client_SSN 							= ""
	first_name 							= ""
	excel_date_notice_sent				= ""
	worker_entered_cleared_status	    = ""
	excel_date_contact_with_other_state = ""
	excel_date_cleared					= ""

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
	FOR i = 1 to 10
		objExcel.Columns(i).AutoFit()
	NEXT
	'add pf3 at the end of the run and error handling for blank cleared status'
	STATS_counter = STATS_counter - 1		'removes 1 to correct the count
	script_end_procedure_with_error_report("Success! The IEVS match cases have now been updated. Please review the NOTES section to review the cases/follow up work to be completed.")
