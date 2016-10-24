'Required for statistical purposes===============================================================================
name_of_script = "BULK - DAIL REPORT.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 30                               'manual run time, per line, in seconds
STATS_denomination = "I"       'I is for each ITEM
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

BeginDialog bulk_dail_report_dialog, 0, 0, 361, 150, "Bulk DAIL report dialog"
  EditBox 10, 35, 345, 15, x_number_editbox
  CheckBox 20, 85, 25, 10, "ALL", All_check
  ButtonGroup ButtonPressed
    OkButton 250, 125, 50, 15
    CancelButton 305, 125, 50, 15
  CheckBox 60, 85, 30, 10, "COLA", cola_check
  CheckBox 105, 85, 30, 10, "CLMS", clms_check
  CheckBox 150, 85, 30, 10, "CSES", cses_check
  CheckBox 195, 85, 30, 10, "ELIG", elig_check
  CheckBox 240, 85, 30, 10, "IEVS", ievs_check
  CheckBox 280, 85, 30, 10, "INFO", info_check
  CheckBox 20, 100, 25, 10, "IV-E", iv3_check
  CheckBox 60, 100, 25, 10, "MA", ma_check
  CheckBox 105, 100, 30, 10, "MEC2", mec2_check
  CheckBox 150, 100, 35, 10, "PARI", pari_chck
  CheckBox 195, 100, 30, 10, "PEPR", pepr_check
  CheckBox 240, 100, 30, 10, "TIKL", tikl_check
  CheckBox 280, 100, 30, 10, "WF1", wf1_check
  Text 145, 5, 90, 10, "---BULK DAIL REPORT---"
  Text 10, 20, 350, 10, "Please enter the x1 numbers of the caseloads you wish to check, separated by commas (if more than one):"
  Text 10, 55, 290, 10, "Note: please enter the entire 7-digit number x1 number. (Example: ''x100abc, x100abc'')"
  GroupBox 5, 70, 305, 50, "Select the type(s) of DAIL message to add to the report:"
EndDialog

'Connects to MAXIS
EMConnect ""

'Looks up an existing user for autofilling the next dialog
CALL find_variable("User: ", x_number_editbox, 7)

'defaulting the script to check all DAILS on a DAIL list
all_check = 1

'Shows the dialog. Doesn't need to loop since we already looked at MAXIS.
DO
	dialog bulk_dail_report_dialog
	cancel_confirmation
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
Loop until are_we_passworded_out = false					'loops until user passwords back in					

'splits the results of the editbox into an array
x_number_array = split(x_number_editbox, ",")

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Changes name of Excel sheet to "DAIL List"
ObjExcel.ActiveSheet.Name = "DAIL List"

'Excel headers and formatting the columns
objExcel.Cells(1, 1).Value = "X1 NUMBER"
objExcel.Cells(1, 1).Font.Bold = True
objExcel.Cells(1, 2).Value = "CASE NBR"
objExcel.Cells(1, 2).Font.Bold = True
objExcel.Cells(1, 3).Value = "CLIENT NAME"
objExcel.Cells(1, 3).Font.Bold = True
objExcel.Cells(1, 4).Value = "DAIL TYPE"
objExcel.Cells(1, 4).Font.Bold = True
objExcel.Cells(1, 5).Value = "DAIL MONTH"
objExcel.Cells(1, 5).Font.Bold = True
objExcel.Cells(1, 6).Value = "DAIL MESSAGE"
objExcel.Cells(1, 6).Font.Bold = True

'Sets variable for all of the Excel stuff
excel_row = 2

'This for...next contains each worker indicated above
For each x_number in x_number_array

	'Trims the x_number so that we don't have glitches
	x_number = trim(x_number)

	back_to_SELF
	MAXIS_case_number = ""			'Blanking this out for PRIV case handling.
	CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
	EMWriteScreen x_number, 21, 6
	transmit
	
	'selecting the type of DAIl message
	EMWriteScreen "x", 4, 12		'transmits to the PICK screen
	transmit
	EMWriteScreen "_", 7, 39		'clears the all selection
	If all_check = 1 then EMWriteScreen "x", 7, 39
	If cola_check = 1 then EMWriteScreen "x", 8, 39
	If clms_check = 1 then EMWriteScreen "x", 9, 39
	If cses_check = 1 then EMWriteScreen "x", 10, 39
	If elig_check = 1 then EMWriteScreen "x", 11, 39
	If ievs_check = 1 then EMWriteScreen "x", 12, 39
	If info_check = 1 then EMWriteScreen "x", 13, 39
	If iv3_check = 1 then EMWriteScreen "x", 14, 39
	If ma_check = 1 then EMWriteScreen "x", 15, 39
 	If mec2_check = 1 then EMWriteScreen "x", 16, 39
	If pari_chck = 1 then EMWriteScreen "x", 17, 39
	If pepr_check = 1 then EMWriteScreen "x", 18, 39
	If tikl_check = 1 then EMWriteScreen "x", 19, 39
	If wf1_check = 1 then EMWriteScreen "x", 20, 39
	transmit
	EMReadScreen number_of_dails, 1, 3, 67		'Reads where the count of DAILs is listed
	DO
		If number_of_dails = " " Then 			'if this space is blank the rest of the DAIL reading is skipped
			objExcel.Cells(excel_row, 1).Value = x_number
			objExcel.Cells(excel_row, 3).Value = "No DAILs for this worker."
			excel_row = excel_row + 1
			Exit Do
		End If 
		'Reading and trimming the MAXIS case number and dumping it in Excel
		EMReadScreen maxis_case_number, 8, 5, 73
		maxis_case_number = trim(maxis_case_number)
		objExcel.Cells(excel_row, 2).Value = maxis_case_number

		'This bit of code grabs the client name. The do/loop expands the search area until the value for
		'next_two equals "--" ... at which time the script determines that the cl name has ended
		dail_col = 6
		name_len = 1
		DO
			EMReadScreen client_name, name_len, 5, 5
			EMReadScreen next_two, 2, 5, dail_col
			IF next_two <> "--" THEN
				name_len = name_len + 1
				dail_col = dail_col + 1
			END IF
		LOOP UNTIL next_two = "--"
		'Dumping the client name in Excel
		objExcel.Cells(excel_row, 3).Value = client_name

		'This is where the script starts reading the DAIL messages.
		'Because the script brings each new case to the top of the page, dail_row starts at 6.
		dail_row = 6
		DO
			'Determining if there is a new case number...
			EMReadScreen new_case, 8, dail_row, 63
			new_case = trim(new_case)
			IF new_case <> "CASE NBR" THEN
				'...if there is NOT a new case number, the script will read the DAIL type, month, year, and message...
				EMReadScreen dail_type, 4, dail_row, 6
				EMReadScreen dail_month, 8, dail_row, 11
				dail_month = trim(dail_month)
				EMReadScreen dail_msg, 61, dail_row, 20
				dail_msg = trim(dail_msg)
				IF dail_msg <> "" AND dail_type <> "    " and dail_month <> "" THEN
					'...and put that in Excel.
					objExcel.Cells(excel_row, 1).Value = x_number
					objExcel.Cells(excel_row, 2).Value = maxis_case_number
					objExcel.Cells(excel_row, 3).Value = client_name
					objExcel.Cells(excel_row, 4).Value = dail_type
					objExcel.Cells(excel_row, 5).Value = dail_month
					objExcel.Cells(excel_row, 6).Value = dail_msg
				END IF

				'...going to the next ding dang row...
				dail_row = dail_row + 1


				'...going to the next page if necessary
				IF dail_row = 19 AND dail_msg <> "" THEN
					PF8
					dail_row = 6
				ELSEIF dail_row = 19 AND dail_msg = "" THEN
					EMReadScreen more_pages, 7, 19, 3
					if more_pages = "More: -" OR more_pages = "       " then
						all_done = true
						'If the script determines that it is on the last page, it EXITS DO...
						exit do
					else
						PF8
						dail_row = 6
					end if
				end if

				if objExcel.Cells(excel_row, 2).value <> "" then
					excel_row = excel_row + 1			'only does this if there's data there (if no data has been entered, it means we're at the end of a DAIL list of some type somehow)
					STATS_counter = STATS_counter + 1 	'adds one instance to the stats counter
				end if
			ELSEIF new_case = "CASE NBR" THEN
				'...if the script does find that there is a new case number (indicated by the presence
				'   of "CASE NBR", it will write a "T" in the next row and transmit, bringing that
				'   case number to the top of your DAIL
				EMWriteScreen "T", dail_row + 1, 3
				transmit
			END IF
		LOOP UNTIL new_case = "CASE NBR" OR (dail_type = "    " AND dail_month = "     " AND dail_msg = "")
		IF all_done = true THEN exit do
	LOOP

	if x_number <> x_number_array(ubound(x_number_array)) then all_done = false
Next

'Enters info about runtime for the benefit of folks using the script
objExcel.Cells(2, 8).Value = "Lines added to Excel sheet:"
objExcel.Cells(3, 8).Value = "Average time to find/select/copy/paste one line (in seconds):"
objExcel.Cells(4, 8).Value = "Estimated manual processing time (lines x average):"
objExcel.Cells(5, 8).Value = "Script run time (in seconds):"
objExcel.Cells(6, 8).Value = "Estimated time savings by using script (in minutes):"
objExcel.Columns(8).Font.Bold = true
objExcel.Cells(2, 9).Value = STATS_counter
objExcel.Cells(3, 9).Value = STATS_manualtime
objExcel.Cells(4, 9).Value = STATS_counter * STATS_manualtime
objExcel.Cells(5, 9).Value = timer - start_time
objExcel.Cells(6, 9).Value = ((STATS_counter * STATS_manualtime) - (timer - start_time)) / 60

'Formatting the column width.
FOR i = 1 to 9
	objExcel.Columns(i).AutoFit()
NEXT

'Going to another sheet, to enter worker-specific statistics
ObjExcel.Worksheets.Add().Name = "DAIL stats by worker"

col_to_use = 3

'Headers
ObjExcel.Cells(1, 2).Value = "DAIL STATS BY WORKER"
ObjExcel.Cells(1, 2).Font.Bold = TRUE
ObjExcel.Cells(2, 1).Value = "WORKER"
objExcel.Cells(2, 1).Font.Bold = TRUE
ObjExcel.Cells(2, 2).Value = "TOTAL"
objExcel.Cells(2, 2).Font.Bold = TRUE

IF all_check = checked OR cola_check = checked THEN
	ObjExcel.Cells(2, col_to_use).Value = "COLA"
	objExcel.Cells(2, col_to_use).Font.Bold = TRUE
	COLA_col = col_to_use
	col_to_use = col_to_use + 1 
	COLA_letter_col = convert_digit_to_excel_column(COLA_col)
END IF
	
IF all_check = checked OR clms_check = checked THEN
	ObjExcel.Cells(2, col_to_use).Value = "CLMS"
	objExcel.Cells(2, col_to_use).Font.Bold = TRUE
	CLMS_col = col_to_use
	col_to_use = col_to_use + 1 
	CLMS_letter_col = convert_digit_to_excel_column(CLMS_col)
END IF

IF all_check = checked OR cses_check = checked THEN
	ObjExcel.Cells(2, col_to_use).Value = "CSES"
	objExcel.Cells(2, col_to_use).Font.Bold = TRUE
	CSES_col = col_to_use
	col_to_use = col_to_use + 1 
	CSES_letter_col = convert_digit_to_excel_column(CSES_col)
END IF
	
IF all_check = checked OR elig_check = checked THEN
	ObjExcel.Cells(2, col_to_use).Value = "ELIG"
	objExcel.Cells(2, col_to_use).Font.Bold = TRUE
	ELIG_col = col_to_use
	col_to_use = col_to_use + 1 
	ELIG_letter_col = convert_digit_to_excel_column(ELIG_col)
END IF
	
IF all_check = checked OR ievs_check = checked THEN
	ObjExcel.Cells(2, col_to_use).Value = "IEVS"
	objExcel.Cells(2, col_to_use).Font.Bold = TRUE
	IEVS_col = col_to_use
	col_to_use = col_to_use + 1 
	IEVS_letter_col = convert_digit_to_excel_column(IEVS_col)
END IF
	
IF all_check = checked OR info_check = checked THEN
	ObjExcel.Cells(2, col_to_use).Value = "INFO"
	objExcel.Cells(2, col_to_use).Font.Bold = TRUE
	INFO_col = col_to_use
	col_to_use = col_to_use + 1 
	INFO_letter_col = convert_digit_to_excel_column(INFO_col)
END IF
	
IF all_check = checked OR iv3_check = checked THEN
	ObjExcel.Cells(2, col_to_use).Value = "IV-E"
	objExcel.Cells(2, col_to_use).Font.Bold = TRUE
	IV3_col = col_to_use
	col_to_use = col_to_use + 1 
	IV3_letter_col = convert_digit_to_excel_column(IV3_col)
END IF
	
IF all_check = checked OR ma_check = checked THEN
	ObjExcel.Cells(2, col_to_use).Value = "MA"
	objExcel.Cells(2, col_to_use).Font.Bold = TRUE
	MA_col = col_to_use
	col_to_use = col_to_use + 1 
	MA_letter_col = convert_digit_to_excel_column(MA_col)
END IF
	
IF all_check = checked OR mec2_check = checked THEN
	ObjExcel.Cells(2, col_to_use).Value = "MEC2"
	objExcel.Cells(2, col_to_use).Font.Bold = TRUE
	MEC2_col = col_to_use
	col_to_use = col_to_use + 1 
	MEC2_letter_col = convert_digit_to_excel_column(MEC2_col)
END IF
	
IF all_check = checked OR pari_chck = checked THEN
	ObjExcel.Cells(2, col_to_use).Value = "PARI"
	objExcel.Cells(2, col_to_use).Font.Bold = TRUE
	PARI_col = col_to_use
	col_to_use = col_to_use + 1 
	PARI_letter_col = convert_digit_to_excel_column(PARI_col)
END IF
	
IF all_check = checked OR pepr_check = checked THEN
	ObjExcel.Cells(2, col_to_use).Value = "PEPR"
	objExcel.Cells(2, col_to_use).Font.Bold = TRUE
	PEPR_col = col_to_use
	col_to_use = col_to_use + 1 
	PEPR_letter_col = convert_digit_to_excel_column(PEPR_col)
END IF
	
IF all_check = checked OR tikl_check = checked THEN
	ObjExcel.Cells(2, col_to_use).Value = "TIKL"
	objExcel.Cells(2, col_to_use).Font.Bold = TRUE
	TIKL_col = col_to_use
	col_to_use = col_to_use + 1 
	TIKL_letter_col = convert_digit_to_excel_column(TIKL_col)
END IF
	
IF all_check = checked OR wf1_check = checked THEN
	ObjExcel.Cells(2, col_to_use).Value = "WF1"
	objExcel.Cells(2, col_to_use).Font.Bold = TRUE
	WF1_col = col_to_use
	col_to_use = col_to_use + 1 
	WF1_letter_col = convert_digit_to_excel_column(WF1_col)
END IF 


'Writes each worker from the worker_array in the Excel spreadsheet
For x = 0 to ubound(x_number_array)
	ObjExcel.Cells(x + 3, 1) = trim(x_number_array(x))
	ObjExcel.Cells(x + 3, 2) = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ")"
	
	'Counts the number of DAILs for each worker based on type and enters it into the correct cell
	IF all_check = checked OR cola_check = checked THEN ObjExcel.Cells(x + 3, COLA_col) = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "COLA" & Chr(34) & ")"	
	IF all_check = checked OR clms_check = checked THEN ObjExcel.Cells(x + 3, CLMS_col) = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "DMND" & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "CRAA" & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "BILL" & Chr(34) & ")"	
	IF all_check = checked OR cses_check = checked THEN ObjExcel.Cells(x + 3, CSES_col) = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "CSES" & Chr(34) & ")"	
	IF all_check = checked OR elig_check = checked THEN ObjExcel.Cells(x + 3, ELIG_col) = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "REIN" & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "STAT" & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "DWP " & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "FS  " & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "CASH" & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "HC  " & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "CCOL" & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "GA  " & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "GRH " & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "MSA " & Chr(34) & ")"	
	IF all_check = checked OR ievs_check = checked THEN ObjExcel.Cells(x + 3, IEVS_col) = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "WAGE" & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "UNVI" & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "BEER" & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "UBEN" & Chr(34) & ")"	
	IF all_check = checked OR info_check = checked THEN ObjExcel.Cells(x + 3, INFO_col) = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "INFO" & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "ISPI" & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "HIRE" & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "SSN " & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "HC" & Chr(34) & ")"	
	IF all_check = checked OR iv3_check  = checked THEN ObjExcel.Cells(x + 3, IV3_col)  = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "IV-E" & Chr(34) & ")"	
	IF all_check = checked OR ma_check   = checked THEN ObjExcel.Cells(x + 3, MA_col)   = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "MA  " & Chr(34) & ")"	
	IF all_check = checked OR mec2_check = checked THEN ObjExcel.Cells(x + 3, MEC2_col) = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "MEC2" & Chr(34) & ")"	
	IF all_check = checked OR pari_chck  = checked THEN ObjExcel.Cells(x + 3, PARI_col) = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "PARI" & Chr(34) & ")"	
	IF all_check = checked OR pepr_check = checked THEN ObjExcel.Cells(x + 3, PEPR_col) = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "PEPR" & Chr(34) & ")"	
	IF all_check = checked OR tikl_check = checked THEN ObjExcel.Cells(x + 3, TIKL_col) = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "TIKL" & Chr(34) & ")"	
	IF all_check = checked OR wf1_check  = checked THEN ObjExcel.Cells(x + 3, WF1_col)  = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & x + 3 & ", 'DAIL List'!D:D, " & Chr(34) & "WF1 " & Chr(34) & ")"	
Next

'Merging header cell.
ObjExcel.Range(ObjExcel.Cells(1, 1), ObjExcel.Cells(1, col_to_use - 1)).Merge

'Centering the cell
objExcel.Cells(1, 2).HorizontalAlignment = -4108

'Autofitting columns
For col_to_autofit = 1 to col_to_use
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! The workers' DAILs are now entered.")
