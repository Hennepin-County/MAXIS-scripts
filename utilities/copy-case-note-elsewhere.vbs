'Required for statistical purposes==========================================================================================
name_of_script = "UTILITIES - COPY CASE NOTE ELSEWHERE.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 15           'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

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
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Checks for county info from global variables, or asks if it is not already defined.
get_county_code

'commented out as WCOM portion isn't being used. If enhanced in the future this code could be valuable.
''THE DIALOG----------------------------------------------------------------------------------------------------
'BeginDialog WCOM_month_dlg, 0, 0, 196, 60, "Enter approval month for WCOM"
'  EditBox 65, 10, 30, 15, approval_month
'  EditBox 155, 10, 30, 15, approval_year
'  ButtonGroup ButtonPressed
'    OkButton 40, 35, 50, 15
'    CancelButton 95, 35, 50, 15
'  Text 100, 15, 55, 10, "Approval Year:"
'  Text 5, 15, 55, 10, "Approval Month:"
'EndDialog


'The script----------------------------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""

'the script will first check similiar to the DAIL SCRUBBER to see if you're in CASE/NOTE and your cursor is resting on a case note you want to use.
EMReadscreen on_case_note_check, 6, 3, 16
IF on_case_note_check <> "Update" THEN script_end_procedure("You are not on CASE/NOTE. Please navigate to CASE/NOTE, leave your cursor on the case note you wish to copy, and re-run the script.")
EMReadscreen MAXIS_case_number, 8, 20, 38
MAXIS_case_number = trim(MAXIS_case_number)
EMGetCursor CN_row, CN_col
IF CN_row < 5 or CN_row > 18 THEN script_end_procedure("Your cursor is not resting on a CASE/NOTE. Please move your cursor and try again. ")

'reading title of case note so it can display on dialog
EMReadscreen current_CN_title, 73, CN_row, CN_col + 3

'This needs to be built AFTER the case note title is read or it won't be able to fill into the text in the dialog.
BeginDialog dialog1, 0, 0, 346, 110, "Copy CASE/NOTE to.."
  EditBox 160, 5, 60, 15, MAXIS_case_number
  DropListBox 165, 65, 65, 15, "Select One..."+chr(9)+"CCOL NOTES"+chr(9)+"SPEC MEMO", copy_location
  ButtonGroup ButtonPressed
    OkButton 115, 85, 50, 15
    CancelButton 170, 85, 50, 15
  Text 105, 10, 55, 10, "Case Number:"
  Text 85, 60, 75, 15, "Select where to copy CASE/NOTE to:"
  Text 5, 35, 50, 10, "Selected Note:"
  Text 65, 30, 280, 25, current_CN_title
EndDialog

'Dialog starts
DO
	err_msg = ""
	Dialog
	cancel_confirmation
	IF MAXIS_case_number = "" THEN err_msg = err_msg & "Please enter a case number" & vbCr
	IF copy_location = "Select One..." THEN err_msg = err_msg & "Please select where to copy this case note" & vbCr
	IF err_msg <> "" THEN msgbox "Please resolve the following:" & vbCr & vbCr & err_msg
Loop until err_msg = ""

Call check_for_MAXIS(False)

'Reading the case note selected
EMWritescreen "x", CN_row, CN_col
Transmit
DO
	note_row = 4								'resetting read row after pf8ing
	DO
		EMReadscreen case_note_contents, 76, note_row, 3
		IF TRIM(case_note_contents) <> "" THEN 														'ignoring the blank lines at the end of a case note
			final_case_note_to_write = final_case_note_to_write & case_note_contents & "|"			'prepping what has been read for the array
		END IF
		note_row = note_row + 1
	LOOP UNTIL note_row = 18
	PF8											'after whole page is read script will try to PF8 to the next page
	page_count = 1 'counting pages read as it changes stats time
	EMReadscreen last_page_check, 4, 24, 14
LOOP UNTIL last_page_check = "LAST"

'Breaking case note up into an array for ease of writing
case_note_array = split(final_case_note_to_write, "|")

'WCOM and MEMOS are limited to 2 smaller pages as such they can only contain 840 (for WCOMS) and 1680(for memos) characters, 15 rows at 56 characters per row per page
'as such script must end if case note that is too big was selected to be entered into MEMO/WCOM.
FOR EACH case_note_line IN case_note_array
	character_count = character_count + len(case_note_line)
NEXT

'Copying to CCOL NOTE
IF copy_location = "CCOL NOTES" THEN
	Call navigate_to_MAXIS_screen("CCOL", "CLSM")
	DO															'getting claim number and making sure it is correct
		DO
			MAXIS_claim_number = InputBox("Claim Number:", "Claim Number")
			IF IsNumeric(MAXIS_claim_number) = FALSE THEN msgbox "Enter a valid claim number"
		LOOP UNTIL IsNumeric(MAXIS_claim_number) = TRUE
		EMWritescreen MAXIS_claim_number, 4, 9
		Transmit
		'Checking to make sure you're in right claim.
		double_check = msgbox("Is this the correct claim?", vbYesNoCancel + vbSystemModal)
		IF double_check = vbCancel THEN stopscript
	LOOP until double_check = vbYes
	PF4
	PF9
	'writing case note into claim note
	FOR EACH case_note_line IN case_note_array
		call write_variable_in_case_note(case_note_line)
	NEXT
END IF

'Copying to SPEC MEMO
IF copy_location = "SPEC MEMO" THEN
	IF character_count > 1680 THEN script_end_procedure("The CASE/NOTE you have selected is too large to copy to a SPEC/MEMO. Please process manually at this time.")
	CALL navigate_to_MAXIS_screen("SPEC", "MEMO")
	'Checking for privileged
	EMReadScreen privileged_case, 40, 24, 2
	IF InStr(privileged_case, "PRIVILEGED") <> 0 THEN
		privileged_array = privileged_array & MAXIS_case_number & "~~~"
	ELSE
		PF5
		'Checking for an AREP. If there's an AREP it'll navigate to STAT/AREP, check to see if the forms go to the AREP. If they do, it'll write X's in those fields below.
		row = 4                             'Defining row and col for the search feature.
		col = 1
		EMSearch "ALTREP", row, col         'Row and col are variables which change from their above declarations if "ALTREP" string is found.
		IF row > 4 THEN                     'If it isn't 4, that means it was found.
			arep_row = row                                          'Logs the row it found the ALTREP string as arep_row
			call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
			EMReadscreen forms_to_arep, 1, 10, 45                   'Reads for the "Forms to AREP?" Y/N response on the panel.
			call navigate_to_MAXIS_screen("SPEC", "MEMO")           'Navigates back to SPEC/MEMO
			PF5                                                     'PF5s again to initiate the new memo process
		END IF
		'Checking for SWKR
		row = 4                             'Defining row and col for the search feature.
		col = 1
		EMSearch "SOCWKR", row, col         'Row and col are variables which change from their above declarations if "SOCWKR" string is found.
		IF row > 4 THEN                     'If it isn't 4, that means it was found.
			swkr_row = row                                          'Logs the row it found the SOCWKR string as swkr_row
			call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
			EMReadscreen forms_to_swkr, 1, 15, 63                'Reads for the "Forms to SWKR?" Y/N response on the panel.
			call navigate_to_MAXIS_screen("SPEC", "MEMO")         'Navigates back to SPEC/MEMO
			PF5                                           'PF5s again to initiate the new memo process
		END IF
		EMWriteScreen "x", 5, 10                                        'Initiates new memo to client
		IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
		IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
		transmit
		FOR EACH case_note_line IN case_note_array						'writing each line into SPEC/MEMO
			CALL write_variable_in_SPEC_MEMO(case_note_line)
		NEXT
	END IF
END IF


'commenting out for time being as in order to maintain the structure of the original CN we have to keep spaces. However this means that spaces are
'used towards the length count. As such a single page of a case note is 1241 characters meaning we could never add a CN to a WCOM. We also cannot ignore blank lines
'as someone may write a blank line as a spacer in their case note. Leaving code in as it may be valuable if this will be an enhancement.
'Copying to SPEC WCOM
'IF copy_location = "SPEC WCOM" THEN
'	IF character_count > 840 THEN script_end_procedure("The CASE/NOTE you have selected is too large to copy to a SPEC/WCOM. Please process manually at this time.")
'	call navigate_to_MAXIS_screen("spec", "wcom")
'
'	DO
'		err_msg = ""
'		dialog WCOM_month_dlg
'		cancel_confirmation
'		IF len(approval_month) <> 2 THEN err_msg = err_msg & "Please enter your month in MM format." & vbNewLine
'		IF len(approval_year) <> 2 THEN err_msg = err_msg & "Please enter your year in YY format." & vbNewLine
'		IF err_msg <> "" THEN msgbox "Please resolve the following:" & vbCr & vbCr & err_msg
'	LOOP until err_msg = ""
'
'	EMWriteScreen approval_month, 3, 46
'	EMWriteScreen approval_year, 3, 51
'	transmit
'
'	DO 								'This DO/LOOP resets to the first page of notices in SPEC/WCOM
'		EMReadScreen more_pages, 8, 18, 72
'		IF more_pages = "MORE:  -" THEN PF7
'	LOOP until more_pages <> "MORE:  -"
'
'	read_row = 7
'	DO
'		waiting_check = ""
'		EMReadscreen prog_type, 2, read_row, 26
'		EMReadscreen waiting_check, 7, read_row, 71 'finds if notice has been printed
'		If waiting_check = "Waiting" and prog_type = "FS" THEN 'checking program type and if it's been printed
'			EMSetcursor read_row, 13
'			EMSendKey "x"
'			Transmit
'			pf9
'			FOR EACH case_note_line IN case_note_array
'				'following code is taken and slightly altered from function write_variable_in_SPEC_MEMO
'				WCOM_col = 15										'The memo col should always be 15 at this point, because it's the beginning. But, this will be dynamically recreated each time.
'				'The following figures out if we need a new page
'				Do
'					EMReadScreen character_test, 1, WCOM_row, WCOM_col 	'Reads a single character at the memo row/col. If there's a character there, it needs to go down a row, and look again until there's nothing. It also needs to trigger these events if it's at or above row 18 (which means we're beyond memo range).
'					If character_test <> " " or WCOM_row >= 18 then
'						WCOM_row = WCOM_row + 1
'
'						'If we get to row 18 (which can't be written to), it will go to the next page of the memo (PF8).
'						If WCOM_row >= 18 then
'							script_end_procedure("The CASE/NOTE you have selected is too large to copy to a SPEC/WCOM. CASE/NOTE was only partially written. Please process manually at this time.")
'							WCOM_row = 3					'Resets this variable to 3
'						End if
'					End if
'				Loop until character_test = " "
'
'				'Each word becomes its own member of the array called variable_array.
'				case_note_line = split(variable, " ")
'
'				For each word in case_note_line
'					'If the length of the word would go past col 74 (you can't write to col 74), it will kick it to the next line
'					If len(word) + WCOM_col > 74 then
'						WCOM_row = WCOM_row + 1
'						WCOM_col = 15
'					End if
'
'					'If we get to row 18 (which can't be written to), it will go to the next page of the memo (PF8).
'					If WCOM_row >= 18 then
'						script_end_procedure("The CASE/NOTE you have selected is too large to copy to a SPEC/WCOM. CASE/NOTE was only partially written. Please process manually at this time.")
'						WCOM_row = 3					'Resets this variable to 3
'					End if
'
'					'Writes the word and a space using EMWriteScreen
'					EMWriteScreen word & " ", WCOM_row, WCOM_col
'
'					'Increases WCOM_col the length of the word + 1 (for the space)
'					WCOM_col = WCOM_col + (len(word) + 1)
'				Next
'
'				'After the array is processed, set the cursor on the following row, in col 15, so that the user can enter in information here (just like writing by hand).
'				EMSetCursor WCOM_row + 1, 15
'			NEXT
'			WCOM_count = WCOM_count + 1
'			exit do
'		ELSE
'			read_row = read_row + 1
'		END IF
'		IF read_row = 18 THEN
'			PF8          'Navigates to the next page of notices.  DO/LOOP until read_row = 18
'			read_row = 7
'		End if
'	LOOP until prog_type = "  "
'	If WCOM_count = 0 THEN  'if no waiting FS notice is found
'		script_end_procedure("No Waiting FS elig results were found in this month for this HH member.")
'	END IF
'END IF

STATS_manualtime = STATS_manualtime + 15 * page_count  'multiplying pages copied from a case note

'Script end procedure
script_end_procedure("")
