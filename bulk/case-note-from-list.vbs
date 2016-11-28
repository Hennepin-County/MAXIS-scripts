'Required for statistical purposes===============================================================================
name_of_script = "BULK - CASE NOTE FROM LIST.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 180                               'manual run time in seconds
STATS_denomination = "C"       'C is for each Case
'END OF stats block==============================================================================================

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

'Dialogs
'>>>>>Main dlg<<<<<
BeginDialog main_menu, 0, 0, 201, 65, "Case Note from List"
  DropListBox 5, 40, 80, 10, "Manual Entry"+chr(9)+"REPT/ACTV"+chr(9)+"Excel File", run_mode
  ButtonGroup ButtonPressed
    OkButton 90, 40, 50, 15
    CancelButton 140, 40, 50, 15
  Text 10, 10, 185, 25, "Please select a run mode for the script. You can either enter the case numbers manually, from REPT/ACTV, or from an Excel file..."
EndDialog

'>>>>> Function to build dlg for manual entry <<<<<
FUNCTION build_manual_entry_dlg(case_number_array, case_note_header, case_note_body, worker_signature)
	'Array for all case numbers
	'This was chosen over building a dlg with 50 variables
	REDim all_cases_array(50, 0)

	BeginDialog man_entry_dlg, 0, 0, 331, 330, "Enter MAXIS case numbers"
		Text 10, 15, 140, 10, "Enter MAXIS case numbers below..."
		dlg_row = 30
		dlg_col = 10
		FOR i = 1 TO 50
			EditBox dlg_col, dlg_row, 55, 15, all_cases_array(i, 0)
			dlg_row = dlg_row + 20
			IF dlg_row = 230 THEN
				dlg_row = 30
				dlg_col = dlg_col + 65
			END IF
		NEXT
		text 10, 235, 120, 10, "Enter case note below"
		Text 10, 255, 25, 10, "Header:"
		Text 10, 275, 20, 10, "Body:"
		Text 10, 295, 60, 10, "Worker Signature:"
		EditBox 45, 250, 280, 15, case_note_header
		EditBox 35, 270, 290, 15, case_note_body
		EditBox 75, 290, 150, 15, worker_signature
		ButtonGroup ButtonPressed
			OkButton 220, 310, 50, 15
			CancelButton 270, 310, 50, 15
	EndDialog

	'Calling the dlg within the function
	DO
		'err_msg handling
		err_msg = ""
		DIALOG man_entry_dlg
			cancel_confirmation
			FOR i = 1 TO 50
				all_cases_array(i, 0) = replace(all_cases_array(i, 0), " ", "")
				IF all_cases_array(i, 0) <> "" THEN
					IF len(all_cases_array(i, 0)) > 8 THEN err_msg = err_msg & vbCr & "* Case number " & all_cases_array(i, 0) & " is too long to be a valid MAXIS case number."
					IF isnumeric(all_cases_array(i, 0)) = FALSE THEN err_msg = err_msg & vbCr & "* Case number " & all_cases_array(i, 0) & " contains alphabetic characters. These are not valid."
				END IF
			NEXT
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	LOOP UNTIL err_msg = ""

	'building the array
	case_number_array = ""
	FOR i = 1 TO 50
		IF all_cases_array(i, 0) <> "" THEN case_number_array = case_number_array & all_cases_array(i, 0) & "~~~"
	NEXT
END FUNCTION

'>>>>>DLG for Excel mode<<<<<
BeginDialog CASE_NOTE_from_excel_dlg, 0, 0, 256, 135, "Case Note Information"
  EditBox 220, 10, 25, 15, excel_col
  EditBox 65, 30, 40, 15, excel_row
  EditBox 190, 30, 40, 15, end_row
  EditBox 45, 50, 205, 15, case_note_header
  EditBox 35, 70, 215, 15, case_note_body
  EditBox 75, 90, 150, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 130, 115, 55, 15
    CancelButton 190, 115, 60, 15
  Text 10, 15, 205, 10, "Please enter the column containing the MAXIS case numbers..."
  Text 10, 35, 50, 10, "Row to start..."
  Text 135, 35, 50, 10, "Row to end..."
  Text 10, 55, 25, 10, "Header:"
  Text 10, 95, 60, 10, "Worker Signature:"
  Text 10, 75, 20, 10, "Body:"
EndDialog

BeginDialog worker_number_dlg, 0, 0, 231, 130, "Enter worker number and Case Note text..."
  EditBox 145, 10, 65, 15, worker_number
  EditBox 45, 50, 180, 15, case_note_header
  EditBox 30, 70, 190, 15, case_note_body
  EditBox 75, 90, 150, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 60, 110, 50, 15
    CancelButton 115, 110, 50, 15
  Text 10, 15, 130, 10, "Please enter the 7-digit worker number:"
  Text 10, 35, 95, 10, "Enter your Case Note text..."
  Text 10, 55, 25, 10, "Header:"
  Text 10, 95, 60, 10, "Worker Signature:"
  Text 10, 75, 20, 10, "Body:"
EndDialog

'----------FUNCTIONS----------
'-----This function needs to be added to the FUNCTIONS FILE-----
'>>>>> This function converts the letter for a number so the script can work with it <<<<<
FUNCTION convert_excel_letter_to_excel_number(excel_col)
	IF isnumeric(excel_col) = FALSE THEN
		alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
		excel_col = ucase(excel_col)
		IF len(excel_col) = 1 THEN
			excel_col = InStr(alphabet, excel_col)
		ELSEIF len(excel_col) = 2 THEN
			excel_col = (26 * InStr(alphabet, left(excel_col, 1))) + (InStr(alphabet, right(excel_col, 1)))
		END IF
	ELSE
		excel_col = CInt(excel_col)
	END IF
END FUNCTION

'The script===========================
EMConnect ""

CALL check_for_MAXIS(true)
copy_case_note = FALSE

'Checking to see if script is being started on an already created case note
EMReadScreen case_note_check, 10, 2, 33
EMReadScreen case_note_list, 10, 2, 2
EMReadScreen mode_check, 1, 20, 9

'If the script is started from a case note the script will ask if this is the note the worker wants to copy
If case_note_check = "Case Notes" AND case_note_list = "          " Then
	If mode_check = "D" or mode_check = "E" Then
		use_existing_note = MsgBox("It appears that you are currently in a case note that has already been written." & vbNewLine & "Would you like to copy this case note into other cases?", vbYesNo + vbQuestion, "Is this the case note?")
	End If
End If

'If it is the note the worker wants to copy, the script will create the message array from reading the case note lines'\
If use_existing_note = vbYes Then
	copy_case_note = TRUE 	'Creating a boolean variable for future use if needed
	note_row = 4			'Beginning of the case notes
	Do 						'Read each line
		EMReadScreen note_line, 77, note_row, 3
		If trim(note_line) = "" Then Exit Do		'Any blank line indicates the end of the case note because there can be no blank lines in a note
		message_array = message_array & note_line & "~%~"		'putting the lines together
		note_row = note_row + 1
		If note_row = 18 then 									'End of a single page of the case note
			EMReadScreen next_page, 7, note_row, 3
			If next_page = "More: +" Then 						'This indicates there is another page of the case note
				PF8												'goes to the next line and resets the row to read'\
				note_row = 4
			End If
		End If
	Loop until next_page = "More:  " OR next_page = "       "	'No more pages
	message_array = message_array & "**Processed in bulk script**"	'Adding the last line of the case note, indicating the note was bulk entered
	message_array = split(message_array, "~%~")					'Creates the array
	case_note_header = message_array (0)						'This defines the variables for the dialog boxes to come
	For message_line = 1 to (UBound(message_array) - 2)
		case_note_body = case_note_body & ", " & trim(message_array(message_line))
	Next
	case_note_body = right(case_note_body, (len(case_note_body) - 2))
	worker_signature = message_array (UBound(message_array) - 1)
End If

'>>>>> loading the main dialog <<<<<
DIALOG main_menu
	IF ButtonPressed = 0 THEN stopscript
	'>>>>> the script has different ways of building case_number_array
	IF run_mode = "Manual Entry" THEN
		CALL build_manual_entry_dlg(case_number_array, case_note_header, case_note_body, worker_signature)

	ELSEIF run_mode = "REPT/ACTV" THEN
		'script_end_procedure("This mode is not yet supported.")
		CALL find_variable("User: ", worker_number, 7)
		DO
			err_msg = ""
			DIALOG worker_number_dlg
				cancel_confirmation
				worker_number = trim(worker_number)
				IF worker_number = "" THEN err_msg = err_msg & vbCr & "* You must enter a worker number."
				IF len(worker_number) <> 7 THEN err_msg = err_msg & vbCr & "* Your worker number must be 7 characters long."
				IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		LOOP UNTIL err_msg = ""

		CALL check_for_MAXIS(false)

		'Checking that case number is blank so as to get a full REPT/ACTV
		CALL find_variable("Case Nbr: ", MAXIS_case_number, 8)
		MAXIS_case_number = replace(MAXIS_case_number, "_", " ")
		MAXIS_case_number = trim(MAXIS_case_number)
		IF MAXIS_case_number <> "" THEN
			back_to_SELF
			EMWriteScreen "________", 18, 43
		END IF
		'Checking that MAXIS is not already in REPT/ACTV so as to get a full REPT/ACTV
		EMReadScreen at_REPT_ACTV, 4, 2, 48
		IF at_REPT_ACTV = "ACTV" THEN back_to_SELF

		CALL navigate_to_MAXIS_screen("REPT", "ACTV")
		CALL write_value_and_transmit(worker_number, 21, 13)
		'Making sure we are at the beginning of REPT/ACTV
		DO
			PF7
			EMReadScreen page_one, 2, 3, 78
			IF isnumeric(page_one) = false then page_one = page_one * 1  'this is future proofing since reading variables keep switching back from numeric and non numeric.
		LOOP UNTIL page_one = 1

		rept_row = 7
		DO
			EMReadScreen MAXIS_case_number, 8, rept_row, 12
			MAXIS_case_number = trim(MAXIS_case_number)
			IF MAXIS_case_number <> "" THEN
				case_number_array = case_number_array & MAXIS_case_number & "~~~"
				rept_row = rept_row + 1
				IF rept_row = 19 THEN
					EMReadScreen next_page_check, 7, 19, 3			'this prevents the script from erroring out if the worker only has one completely full page of cases.
					If next_page_check = "More: +" Then
						rept_row = 7
						PF8
					Else
						Exit Do
					End If
				END IF
			END IF
		LOOP until MAXIS_case_number = ""

	ELSEIF run_mode = "Excel File" THEN
		'Opening the Excel file

		DO
			call file_selection_system_dialog(excel_file_path, ".xlsx")

			Set objExcel = CreateObject("Excel.Application")
			Set objWorkbook = objExcel.Workbooks.Open(excel_file_path)
			objExcel.Visible = True
			objExcel.DisplayAlerts = True

			confirm_file = MsgBox("Is this the correct file? Press YES to continue. Press NO to try again. Press CANCEL to stop the script.", vbYesNoCancel)
			IF confirm_file = vbCancel THEN
				objWorkbook.Close
				objExcel.Quit
				stopscript
			ELSEIF confirm_file = vbNo THEN
				objWorkbook.Close
				objExcel.Quit
			END IF
		LOOP UNTIL confirm_file = vbYes

		'Gathering the information from the user about the fields in Excel to look for.
		DO
			err_msg = ""
			DIALOG CASE_NOTE_from_excel_dlg
				IF ButtonPressed = 0 THEN stopscript
				IF isnumeric(excel_col) = FALSE AND len(excel_col) > 2 THEN
					err_msg = err_msg & vbCr & "* Please do not use such a large column. The script cannot handle it."
				ELSE
					IF (isnumeric(right(excel_col, 1)) = TRUE AND isnumeric(left(excel_col, 1)) = FALSE) OR (isnumeric(right(excel_col, 1)) = FALSE AND isnumeric(left(excel_col, 1)) = TRUE) THEN
						err_msg = err_msg & vbCr & "* Please use a valid Column indicator. " & excel_col & " contains BOTH a letter and a number."
					ELSE
						call convert_excel_letter_to_excel_number(excel_col)
						IF isnumeric(excel_row) = false or isnumeric(end_row) = false THEN err_msg = err_msg & vbCr & "* Please enter the Excel rows as numeric characters."
						IF end_row = "" THEN err_msg = err_msg & vbCr & "* Please enter an end to the search. The script needs to know when to stop searching."
					END IF
				END IF
				IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		LOOP UNTIL err_msg = ""

		CALL check_for_MAXIS(false)
		'Generating a CASE NOTE for each case.
		FOR i = excel_row TO end_row
			IF objExcel.Cells(i, excel_col).Value <> "" THEN
				case_number_array = case_number_array & objExcel.Cells(i, excel_col).Value & "~~~"
			END IF
		NEXT
	END IF

CALL check_for_MAXIS(false)

'The business of sending Case notes
case_number_array = trim(case_number_array)
case_number_array = split(case_number_array, "~~~")

'Formatting case note
If copy_case_note = FALSE Then
	message_array = case_note_header & "~%~" & case_note_body & "~%~" & "---" & "~%~" & worker_signature & "~%~" & "---" & "~%~" & "**Processed in bulk script**"
	message_array = split(message_array, "~%~")
End If

privileged_array = ""

FOR EACH MAXIS_case_number IN case_number_array
	IF MAXIS_case_number <> "" THEN
		CALL navigate_to_MAXIS_screen("CASE", "NOTE")
		'Checking for privileged
		EMReadScreen privileged_case, 40, 24, 2
		IF InStr(privileged_case, "PRIVILEGED") <> 0 THEN
			privileged_array = privileged_array & MAXIS_case_number & "~~~"
		ELSE
			PF9
			'-----Added because the script was only case noting the header, footer and worker_signature on the first case.
			FOR EACH message_part IN message_array
				CALL write_variable_in_CASE_NOTE(message_part)
				STATS_counter = STATS_counter + 1    'adds one instance to the stats counter
			NEXT
		END IF
	END IF
NEXT

IF privileged_array <> "" THEN
	privileged_array = replace(privileged_array, "~~~", vbCr)
	MsgBox "The script could not generate a CASE NOTE for the following cases..." & vbCr & privileged_array
END IF

STATS_counter = STATS_counter - 1  'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success!!")
