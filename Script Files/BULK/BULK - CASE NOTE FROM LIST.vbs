'Gathering stats==============================================================================
name_of_script = "BULK - CASE NOTE FROM LIST.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
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

'Required for statistical purposes==========================================================================================
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 180                               'manual run time in seconds
STATS_denomination = "C"       'C is for each Case
'END OF stats block==============================================================================================

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

'-------THIS FUNCTION ALLOWS THE USER TO PICK AN EXCEL FILE---------
Function BrowseForFile()
    Dim shell : Set shell = CreateObject("Shell.Application")
    Dim file : Set file = shell.BrowseForFolder(0, "Choose a file:", &H4000, "Computer")
	IF file is Nothing THEN 
		script_end_procedure("The script will end.")
	ELSE
		BrowseForFile = file.self.Path
	END IF
End Function

'The script===========================
EMConnect ""

CALL check_for_MAXIS(true)

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
		CALL find_variable("Case Nbr: ", case_number, 8)
		case_number = replace(case_number, "_", " ")
		case_number = trim(case_number)
		IF case_number <> "" THEN 
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
			EMReadScreen case_number, 8, rept_row, 12
			case_number = trim(case_number)
			IF case_number <> "" THEN 
				case_number_array = case_number_array & case_number & "~~~"
				rept_row = rept_row + 1
				IF rept_row = 19 THEN 
					rept_row = 7 
					PF8
					EMReadScreen last_page_check, 4, 24, 14			'this prevents the script from erroring out if the worker only has one completely full page of cases. 
					If last_page_check = "LAST" THEN EXIT DO
				END IF
			ELSE
				EXIT DO
			END IF
		LOOP 

	ELSEIF run_mode = "Excel File" THEN 
		'Opening the Excel file
		
		DO
			'file_location = InputBox("Please enter the file location.")
			
			Set objExcel = CreateObject("Excel.Application")
			Set objWorkbook = objExcel.Workbooks.Open(BrowseForFile)
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
message_array = case_note_header & "~%~" & case_note_body & "~%~" & "---" & "~%~" & worker_signature & "~%~" & "---" & "~%~" & "**Processed in bulk script**"
message_array = split(message_array, "~%~")

privileged_array = ""

FOR EACH case_number IN case_number_array
	IF case_number <> "" THEN 
		CALL navigate_to_MAXIS_screen("CASE", "NOTE")
		'Checking for privileged
		EMReadScreen privileged_case, 40, 24, 2
		IF InStr(privileged_case, "PRIVILEGED") <> 0 THEN 
			privileged_array = privileged_array & case_number & "~~~"
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
