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
call changelog_update("03/19/2020", "Added Client Contact Follow Up Only option to support assignments for follow up generic phones work. Also updated how DAIL's are read and captured.", "Ilse Ferris, Hennepin County")
call changelog_update("01/28/2019", "Added functionality to remove '=' from any TIKL messages. The equal sign is not able to be written into Excel.", "Ilse Ferris, Hennepin County")
call changelog_update("01/28/2019", "Removed text in spreadsheet that indicates if there is no DAIL for a particular x number. Stats will still relfect the number of DAILS found.", "Ilse Ferris, Hennepin County")
call changelog_update("12/13/2018", "Updated option selection handling, and other background functionality.", "Ilse Ferris, Hennepin County")
call changelog_update("03/12/2018", "Fixed bug for cases with more than one page of DAILs for the same case. Added all agency handling.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
EMConnect ""

'defaulting the script to check all DAILS on a DAIL list
all_check = 1
all_workers_check = 1

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 361, 155, "Bulk DAIL report dialog"
  EditBox 10, 35, 345, 15, worker_number
  CheckBox 40, 85, 25, 10, "ALL", All_check
  CheckBox 80, 85, 30, 10, "COLA", cola_check
  CheckBox 125, 85, 30, 10, "CLMS", clms_check
  CheckBox 170, 85, 30, 10, "CSES", cses_check
  CheckBox 215, 85, 30, 10, "ELIG", elig_check
  CheckBox 260, 85, 30, 10, "IEVS", ievs_check
  CheckBox 300, 85, 30, 10, "INFO", info_check
  CheckBox 40, 100, 25, 10, "IV-E", iv3_check
  CheckBox 80, 100, 25, 10, "MA", ma_check
  CheckBox 125, 100, 30, 10, "MEC2", mec2_check
  CheckBox 170, 100, 35, 10, "PARI", pari_chck
  CheckBox 215, 100, 30, 10, "PEPR", pepr_check
  CheckBox 260, 100, 30, 10, "TIKL", tikl_check
  CheckBox 300, 100, 30, 10, "WF1", wf1_check
  CheckBox 40, 115, 180, 10, "Check here for Client Contact Follow up TIKL's only.", TIKL_FollowUp_checkbox
  CheckBox 10, 140, 135, 10, "Check here to process for all workers.", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 250, 135, 50, 15
    CancelButton 305, 135, 50, 15
  Text 10, 55, 290, 10, "Note: please enter the entire 7-digit number x1 number. (Example: ''x100abc, x100abc'')"
  GroupBox 5, 70, 350, 60, "Select the type(s) of DAIL message to add to the report:"
  Text 145, 5, 90, 10, "---BULK DAIL REPORT---"
  Text 10, 20, 350, 10, "Please enter the x1 numbers of the caseloads you wish to check, separated by commas (if more than one):"
EndDialog

'Shows the dialog. Doesn't need to loop since we already looked at MAXIS.
DO
	Do
        err_msg = ""
        dialog Dialog1
	    Cancel_without_confirmation
	    If trim(worker_number) = "" and all_workers_check = 0 then err_msg = err_msg & vbNewLine & "* Select a worker number(s) or all cases."
	    If trim(worker_number) <> "" and all_workers_check = 1 then err_msg = err_msg & vbNewLine & "* Select a worker number(s) or all cases, not both options."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    LOOP until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else
	x1s_from_dialog = split(worker_number, ",")	'Splits the worker array based on commas

	'Need to add the worker_county_code to each one
	For each x1_number in x1s_from_dialog
		If worker_array = "" then
			worker_array = trim(ucase(x1_number))		'replaces worker_county_code if found in the typed x1 number
		Else
			worker_array = worker_array & ", " & trim(ucase(x1_number)) 'replaces worker_county_code if found in the typed x1 number
		End if
	Next
	'Split worker_array
	worker_array = split(worker_array, ", ")
End if

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Changes name of Excel sheet to "DAIL List"
ObjExcel.ActiveSheet.Name = "DAIL List"

'Excel headers and formatting the columns
objExcel.Cells(1, 1).Value = "X1 NUMBER"
objExcel.Cells(1, 2).Value = "CASE NBR"
objExcel.Cells(1, 3).Value = "CLIENT NAME"
objExcel.Cells(1, 4).Value = "DAIL TYPE"
objExcel.Cells(1, 5).Value = "DAIL MONTH"
objExcel.Cells(1, 6).Value = "DAIL MESSAGE"

FOR i = 1 to 6		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Sets variable for all of the Excel stuff
excel_row = 2

back_to_self
CALL navigate_to_MAXIS_screen("DAIL", "DAIL")

'This for...next contains each worker indicated above
For each worker in worker_array
	DO
		EMReadScreen dail_check, 4, 2, 48
		If next_dail_check <> "DAIL" then
			MAXIS_case_number = ""
			CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
		End if
	Loop until dail_check = "DAIL"

	EMWriteScreen worker, 21, 6
	transmit
	transmit 'transmit past 'not your dail message'

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
    If TIKL_FollowUp_checkbox = 1 then EMWriteScreen "x", 19, 39
	If wf1_check = 1 then EMWriteScreen "x", 20, 39
	transmit

    EMReadScreen number_of_dails, 1, 3, 67		'Reads where the count of DAILs is listed

	DO
		If number_of_dails = " " Then exit do		'if this space is blank the rest of the DAIL reading is skipped

		dail_row = 6			'Because the script brings each new case to the top of the page, dail_row starts at 6.
		DO
			dail_type = ""
			dail_msg = ""

		    'Determining if there is a new case number...
		    EMReadScreen new_case, 8, dail_row, 63
		    new_case = trim(new_case)
		    IF new_case <> "CASE NBR" THEN '...if there is NOT a new case number, the script will read the DAIL type, month, year, and message...
				Call write_value_and_transmit("T", dail_row, 3)
				dail_row = 6
			ELSEIF new_case = "CASE NBR" THEN
			    '...if the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
			    Call write_value_and_transmit("T", dail_row + 1, 3)
				dail_row = 6
			End if

		    '----------------------------------------------------------------------------------------------------CLIENT NAME
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

            EMReadScreen maxis_case_number, 8, dail_row - 1, 73
            EMReadScreen dail_month, 8, dail_row, 11
			EMReadScreen dail_type, 4, dail_row, 6
			EMReadScreen dail_msg, 61, dail_row, 20
            dail_msg = trim(dail_msg)
            If right(dail_msg, 1) = "*" THEN dail_msg = left(dail_msg, len(dail_msg) - 1)
            dail_msg = trim(dail_msg)

			IF trim(dail_msg) <> "" AND dail_type <> "    " and trim(dail_month) <> "" THEN
                If TIKL_FollowUp_checkbox = 1 then
                    If instr(dail_msg, "!!PHONE CONTACT FOLLOW UP REQUIRED!!") then
                        capture_msg = True
                    else
                        capture_msg = False
                    end if
                Else
                    capture_msg = true
                End if

                If capture_msg = True then
				    '...and put that in Excel.
				    objExcel.Cells(excel_row, 1).Value = worker
				    objExcel.Cells(excel_row, 2).Value = maxis_case_number
				    objExcel.Cells(excel_row, 3).Value = client_name
				    objExcel.Cells(excel_row, 4).Value = dail_type
				    objExcel.Cells(excel_row, 5).Value = trim(dail_month)
				    objExcel.Cells(excel_row, 6).Value = trim(dail_msg)
				    excel_row = excel_row + 1			'only does this if there's data there (if no data has been entered, it means we're at the end of a DAIL list of some type somehow)
				    STATS_counter = STATS_counter + 1 	'adds one instance to the stats counter
                End if
			END IF

			'...going to the next ding dang row...
			dail_row = dail_row + 1

            EMReadScreen message_error, 11, 24, 2		'Cases can also NAT out for whatever reason if the no messages instruction comes up.
            If message_error = "NO MESSAGES" then
                CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
                Call write_value_and_transmit(worker, 21, 6)
                transmit   'transmit past 'not your dail message'
                Call dail_selection
                exit do
            End if

            '...going to the next page if necessary
            EMReadScreen next_dail_check, 4, dail_row, 4
            If trim(next_dail_check) = "" then
                PF8
                EMReadScreen last_page_check, 21, 24, 2
                If last_page_check = "THIS IS THE LAST PAGE" then
                    all_done = true
                    exit do
                Else
                    dail_row = 6
                End if
            End if
        LOOP
        IF all_done = true THEN exit do
    LOOP
Next

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)

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
excel_row = 3 'setting the first excel row for stats
For x = 0 to ubound(worker_array)
	ObjExcel.Cells(excel_row, 1) = trim(worker_array(x))
	ObjExcel.Cells(excel_row, 2) = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ")"

	'Counts the number of DAILs for each worker based on type and enters it into the correct cell
	IF all_check = checked OR cola_check = checked THEN ObjExcel.Cells(excel_row, COLA_col) = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "COLA" & Chr(34) & ")"
	IF all_check = checked OR clms_check = checked THEN ObjExcel.Cells(excel_row, CLMS_col) = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "DMND" & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "CRAA" & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "BILL" & Chr(34) & ")"
	IF all_check = checked OR cses_check = checked THEN ObjExcel.Cells(excel_row, CSES_col) = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "CSES" & Chr(34) & ")"
	IF all_check = checked OR elig_check = checked THEN ObjExcel.Cells(excel_row, ELIG_col) = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "REIN" & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "STAT" & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "DWP " & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "FS  " & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "CASH" & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "HC  " & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "CCOL" & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "GA  " & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "GRH " & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "MSA " & Chr(34) & ")"
	IF all_check = checked OR ievs_check = checked THEN ObjExcel.Cells(excel_row, IEVS_col) = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "WAGE" & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "UNVI" & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "BEER" & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "UBEN" & Chr(34) & ")"
	IF all_check = checked OR info_check = checked THEN ObjExcel.Cells(excel_row, INFO_col) = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "INFO" & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "ISPI" & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "HIRE" & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "SSN " & Chr(34) & ") + COUNTIFS('DAIL List'!A:A, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "HC" & Chr(34) & ")"
	IF all_check = checked OR iv3_check  = checked THEN ObjExcel.Cells(excel_row, IV3_col)  = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "IV-E" & Chr(34) & ")"
	IF all_check = checked OR ma_check   = checked THEN ObjExcel.Cells(excel_row, MA_col)   = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "MA  " & Chr(34) & ")"
	IF all_check = checked OR mec2_check = checked THEN ObjExcel.Cells(excel_row, MEC2_col) = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "MEC2" & Chr(34) & ")"
	IF all_check = checked OR pari_chck  = checked THEN ObjExcel.Cells(excel_row, PARI_col) = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "PARI" & Chr(34) & ")"
	IF all_check = checked OR pepr_check = checked THEN ObjExcel.Cells(excel_row, PEPR_col) = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "PEPR" & Chr(34) & ")"
	IF all_check = checked OR tikl_check = checked THEN ObjExcel.Cells(excel_row, TIKL_col) = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "TIKL" & Chr(34) & ")"
	IF all_check = checked OR wf1_check  = checked THEN ObjExcel.Cells(excel_row, WF1_col)  = "=COUNTIFS('DAIL List'!B:B, " & Chr(34) & "<>" & Chr(34) & " & " & Chr(34) & Chr(34) & ", 'DAIL List'!A:A, A" & excel_row & ", 'DAIL List'!D:D, " & Chr(34) & "WF1 " & Chr(34) & ")"
	excel_row = excel_row + 1	'incremenbting to the next excel row for the next list of stats'
Next

'Merging header cell.
ObjExcel.Range(ObjExcel.Cells(1, 1), ObjExcel.Cells(1, col_to_use - 1)).Merge

'Centering the cell
objExcel.Cells(1, 2).HorizontalAlignment = -4108

'Autofitting columns
For col_to_autofit = 1 to col_to_use
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

script_end_procedure("Success! The workers' DAILs are now entered.")
