'Required for statistical purposes===============================================================================
name_of_script = "BULK - UPDATE EOMC LIST.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 20                	'manual run time in seconds
STATS_denomination = "I"       			'I is for each Item
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
call changelog_update("12/28/2016", "Fixed bug with cancel handling for initial dialog.", "David Courtright, Saint Louis County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'this function converts excel column letters to numeric values'
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

'------------------'

opening_msg = MsgBox("This script will check all cases in a saved REPT/EOMC excel file and update the file with current case status.  Press ok to select the saved file to check.  NOTE: The file must maintain the original formatting as created by the REPT/EOMC Bulk script.", vbOkCancel)
IF opening_msg = vbcancel then stopscript

DO 'THIS loop makes sure this is a valid file created by EOMC'
	DO 'This loop opens the file browser and asks user to confirm'

		call file_selection_system_dialog(excel_file_path, ".xlsx")	'Selects an excel file, adds it to excel_file_path

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

		IF objExcel.cells(1, 4).value <> "AUTOCLOSE?" THEN
			confirm_file = MsgBox("This does not appear to be a valid EOMC spreadsheet.  Press OK to try again or cancel to stop the script.", vbOkCancel)
			IF confirm_file = vbCancel THEN
				objWorkbook.Close
				objExcel.Quit
				stopscript
			ELSEIF confirm_file = vbOk THEN
				objWorkbook.Close
				ObjExcel.Quit
			END IF
		END IF
LOOP UNTIL confirm_file <> vbOK

check_for_maxis(true)



'Finding which columns are on the sheet
col_to_use = 5
IF objExcel.cells(1, col_to_use).value = "SNAP?" THEN
	FS_col = col_to_use
	col_to_use = col_to_use + 1
END IF

IF objExcel.cells(1, col_to_use).value = "CASH?" THEN
	cash_col = col_to_use
	col_to_use = col_to_use + 1
END If

IF objExcel.cells(1, col_to_use).value = "HC?" THEN
	HC_col = col_to_use
	col_to_use = col_to_use + 1
END If

IF objExcel.cells(1, col_to_use).value = "EGA?" THEN
	EGA_col = col_to_use
	col_to_use = col_to_use + 1
END If

IF objExcel.cells(1, col_to_use).value = "GRH?" THEN
	GRH_col = col_to_use
	col_to_use = col_to_use + 1
END If
'making sure this is a useable excel file, if col_to_use still equals 5, the REPT/EOMC column headers weren't found.
IF col_to_use = 5 THEN script_end_procedure("This does not appear to be a valid REPT/EOMC file.  The script will now stop.")

'Adding columns for the current status
Const xlShiftToLeft = -4159	'This constant is needed for inserting
col_offset = 0 'this variable will be used to count new cols inserted to make sure we are in the right place'
'SNAP 1st
IF FS_col <> "" THEN
	FS_col_letter = convert_digit_to_excel_column(FS_col + 1 + col_offset) & "1" 'converting the numeric column to a letter/number cell name
	Set	objRange = objExcel.Range(FS_col_letter).EntireColumn 'define the range we want to move'
	objRange.Insert(xlShiftToLeft) 'move it and insert a new column
  FS_col = FS_col + col_offset
	objExcel.cells(1, FS_col + 1).value = "FS Status"
	col_offset = col_offset + 1
END IF

IF cash_col <> "" THEN
	cash_col_letter = convert_digit_to_excel_column(cash_col + 1 + col_offset) & "1" 'converting the numeric column to a letter/number cell name
	Set	objRange = objExcel.Range(cash_col_letter).EntireColumn 'define the range we want to move'
	objRange.Insert(xlShiftToLeft) 'move it and insert a new column
	cash_col = cash_col + col_offset 'assign the new location to the column'
	objExcel.cells(1, cash_col + 1).value = "Cash Status"
	col_offset = col_offset + 1
END IF

IF HC_col <> "" THEN
	HC_col_letter = convert_digit_to_excel_column(HC_col + 1 + col_offset) & "1" 'converting the numeric column to a letter/number cell name
	Set	objRange = objExcel.Range(HC_col_letter).EntireColumn 'define the range we want to move'
	objRange.Insert(xlShiftToLeft) 'move it and insert a new column
	HC_col = HC_col + col_offset
	objExcel.cells(1, HC_col + 1).value = "HC Status"
	col_offset = col_offset + 1
END IF

IF EGA_col <> "" THEN
	EGA_col_letter = convert_digit_to_excel_column(EGA_col + 1 + col_offset) & "1" 'converting the numeric column to a letter/number cell name
	Set	objRange = objExcel.Range(EGA_col_letter).EntireColumn 'define the range we want to move'
	objRange.Insert(xlShiftToLeft) 'move it and insert a new column
	EGA_col = EGA_col + col_offset
	objExcel.cells(1, EGA_col + 1).value = "EGA Status"
	col_offset = col_offset + 1
END IF

IF GRH_col <> "" THEN
	GRH_col_letter = convert_digit_to_excel_column(GRH_col + 1 + col_offset) & "1" 'converting the numeric column to a letter/number cell name
	Set	objRange = objExcel.Range(GRH_col_letter).EntireColumn 'define the range we want to move'
	objRange.Insert(xlShiftToLeft) 'move it and insert a new column
	GRH_col = GRH_col + col_offset
	objExcel.cells(1, GRH_col + 1).value = "GRH Status"
	col_offset = col_offset + 1
END IF

'Going to the first case to begin reading information
excel_row = 2
Do
	MAXIS_case_number = objExcel.Cells(excel_row, 2).value
	call navigate_to_MAXIS_screen("CASE", "CURR")
	'checking for each prog on the listed
	IF fs_col <> "" THEN
		IF objExcel.cells(excel_row, fs_col).value <> "" THEN 'Checking SNAP status
			call find_variable("FS: ", fs_status, 6)
			IF fs_status <> "" THEN ObjExcel.Cells(excel_row, fs_col+1).Value = fs_status
		END If
	END IF
	IF cash_col <> "" THEN
			IF left(objExcel.cells(excel_row, cash_col).value, 2) <> "" THEN 'This checks for generic cash pending'
			call find_variable("Cash: ", cash_status, 6)
			IF cash_status <> "" THEN ObjExcel.Cells(excel_row, cash_col+1).value = cash_status
		END IF
		IF cash_status <> "PENDIN" THEN 'generic pending not found, search each program'
			IF left(objExcel.cells(excel_row, cash_col).value, 2) = "MF" THEN 'checking MFIP status'
				call find_variable("MFIP: ", cash_status, 6)
			ELSEIF left(objExcel.cells(excel_row, cash_col).value, 2) = "MS" THEN 'checking MSA status'
				call find_variable("MSA: ", cash_status, 6)
			ELSEIF left(objExcel.cells(excel_row, cash_col).value, 2) = "GA" THEN 'checking GA status'
				call find_variable("GA: ", cash_status, 6)
			ELSEIF left(objExcel.cells(excel_row, cash_col).value, 2) = "DW" THEN 'checking DWP status'
				call find_variable("DWP: ", cash_status, 6)
			END IF
			IF cash_status <> "" THEN ObjExcel.Cells(excel_row, cash_col+1).value = cash_status
		END If
	END IF
	IF HC_col <> "" THEN
		IF left(objExcel.cells(excel_row, HC_col).value, 2) <> "" THEN 'checking HC status'
			call find_variable("HC: ", HC_status, 6)
			IF HC_status <> "" THEN ObjExcel.Cells(excel_row, HC_col+1).Value = HC_status
		END If
	END IF
	IF GRH_col <> "" THEN
		IF left(objExcel.cells(excel_row, GRH_col).value, 2) <> "" THEN 'checking GRH status'
			call find_variable("GRH: ", GRH_status, 6)
			IF GRH_status <> "" THEN ObjExcel.Cells(excel_row, GRH_col+1).Value = GRH_status
		END If
	END IF
	'reset variables
	fs_status = ""
	cash_status = ""
	HC_status = ""
	GRH_status = ""
	excel_row = excel_row + 1
Loop until MAXIS_case_number = ""

'Autofitting columns
For col_to_autofit = 1 to col_to_use + col_offset
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

STATS_counter = STATS_counter - 1
script_end_procedure("Success. The spreadsheet has been updated with current program status.")
