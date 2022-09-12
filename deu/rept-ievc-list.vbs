'Required for statistical purposes===============================================================================
name_of_script = "BULK - REPT-IEVC LIST.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 300                               'manual run time, per line, in seconds
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
CALL changelog_update("09/12/2022", "Removed VGO verbiage.", "MiKayla Handley, Hennepin County")
CALL changelog_update("06/21/2018", "Updated with requested enhancements.", "MiKayla Handley, Hennepin County")
CALL changelog_update("05/18/2018", "Updated coordinates for writing stats in excel.", "MiKayla Handley, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'Connects to MAXIS
EMConnect ""

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 296, 60, "Pull Income Verifications To Do (IEVC) data into Excel dialog"
  ButtonGroup ButtonPressed
    OkButton 190, 40, 45, 15
    CancelButton 245, 40, 45, 15
  Text 5, 20, 290, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. Do not type when the script is writing to excel. "
  Text 5, 5, 125, 10, "***PULL REPT DATA INTO EXCEL***"
EndDialog

DO
    DO
    	err_msg = ""
    	Dialog Dialog1
    	cancel_without_confirmation
    	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    LOOP UNTIL err_msg = ""
	CALL check_for_password_without_transmit(are_we_passworded_out)
Loop until are_we_passworded_out = false
'Starting the query start time (for the query runtime at the end)
query_start_time = timer
Call check_for_MAXIS(False)
back_to_SELF
Call navigate_to_MAXIS_screen("REPT", "IEVC")

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Name for the current sheet'
ObjExcel.ActiveSheet.Name = "Case information"

'Excel headers and formatting the columns
'------------------------------------------------------IEVC'
objExcel.Cells(1, 1).Value     = "X1 NUMBER" 'x_number
objExcel.Cells(1, 2).Value     = "CASE NUMBER" 'maxis_case_number
objExcel.Cells(1, 3).Value     = "CLIENT NAME" 'client_name
objExcel.Cells(1, 4).Value     = "EARNER NAME" 'client_name
objExcel.Cells(1, 5).Value     = "SSN" 'client_ssn IDLA
objExcel.Cells(1, 6).Value     = "REL"  'client_rel
objExcel.Cells(1, 7).Value     = "TYPE" 'match_type IDLA
objExcel.Cells(1, 8).Value     = "COVERED PERIOD" 'covered_period
objExcel.Cells(1, 9).Value     = "DAYS REMAINING" 'days_remaining
objExcel.Cells(1, 10).Value     = "DOB" 'client_dob
objExcel.Cells(1, 11).Value    = "STATUS" 'overdue
objExcel.Cells(1, 12).Value    = "PROGRAM" 'active_programs
objExcel.Cells(1, 13).Value    = "DIFF NOTICE SENT" 'diff_notc_sent
objExcel.Cells(1, 14).Value    = "DATE DIFF NOTICE SENT" 'diff_notc_date
objExcel.Cells(1, 15).Value    = "AMOUNT" 'income_amount
objExcel.Cells(1, 16).Value    = "YEAR" 'match_year
objExcel.Cells(1, 17).Value    = "EMPLOYER NAME" 'income_source
objExcel.Cells(1, 18).Value    = "NONWAGE INCOME DATE" 'nonwage_date
'objExcel.Cells(1, 19).Value    = "SUPERVISOR ID" 'supervisor_id
'objExcel.Cells(1, 20).Value    = "WORKER NAME" 'worker_name

For excel_row = 1 to 19
	objExcel.Cells(excel_row).Font.Bold = True
Next
'This bit freezes the top row of the Excel sheet for better useability when there is a lot of information
ObjExcel.ActiveSheet.Range("A2").Select
objExcel.ActiveWindow.FreezePanes = True

'Sets variable for all of the Excel stuff
excel_row = 2

CALL clear_line_of_text(4, 11) 'remove worker number
CALL clear_line_of_text(4, 32) 'remove HSS number
TRANSMIT 'clearing the lines and pulling all of the REPT'

EMReadScreen non_disclosure_screen, 14, 2, 46	'Checks to make sure the NDA is current
If non_disclosure_screen = "Non-disclosure" Then script_end_procedure ("It appears you need to confirm agreement to access IEVC. Please navigate there manually to confirm and then run the script again.")

IEVC_Row = 8
'For each ievs_match in ievs_match_array
DO
	EMReadScreen IEVS_message, 7, IEVC_Row, 5

	EMReadScreen x_number, 7, IEVC_Row, 5			'Reads the x number and adds to excel
	objExcel.Cells(excel_row, 1).Value = trim(x_number)		'enters the worker number to the excel spreadsheet

	EMReadScreen maxis_case_number, 8, IEVC_Row, 31		'enters the case number to the excel spreadsheet
	maxis_case_number = trim(maxis_case_number)
	If maxis_case_number = "" then
        IEVC_Row = IEVC_Row + 1
    Else
		objExcel.Cells(excel_row, 2).Value = maxis_case_number	'Adds case number to Excel

	    EMReadScreen client_rel, 02, IEVC_Row, 41			'Reads the client name and adds to excel
	    objExcel.Cells(excel_row, 6).Value = trim(client_rel)

	    EMReadScreen match_type, 3, IEVC_Row, 57			'Reads the client name and adds to excel
	    objExcel.Cells(excel_row, 7).Value = trim(match_type)

	    EMReadScreen client_dob, 10, IEVC_Row, 45			'Reads the client name and adds to excel
	    objExcel.Cells(excel_row, 10).Value = trim(client_dob)

	    EMReadScreen covered_period, 11, IEVC_Row, 62		'Reads the dates of the match and adds to excel
	    objExcel.Cells(excel_row, 8).Value = trim(covered_period)

	    EMReadScreen days_remaining, 6, IEVC_Row, 74		'Reads how the days left to resolve the match and adds to excel
	    days_remaining = trim(days_remaining)
	    objExcel.Cells(excel_row, 9).Value = days_remaining
	    objExcel.Cells(excel_row, 9).NumberFormat = "0"
	    If left(days_remaining, 1) = "(" Then 				'If this is a negative number - listed in () on the panel
	    	objExcel.Cells(excel_row, 11).Value = "OVERDUE!"		'Adds this to the spreadsheet
	    	objExcel.Cells(excel_row, 11).Font.Bold = True 		'Highlights the overdue word
	    	For col = 1 to 19
	    		objExcel.Cells(excel_row, col).Interior.ColorIndex = 3	'Fills the row with red
	    	Next
	    End If

    	EMWriteScreen "D", IEVC_Row, 3		'Opens the detail on the match
	    transmit
	    row = 1
	    col = 1

	    EMReadScreen client_name, 35, 5, 25			'Reads the client name and adds to excel
	    client_name = trim(client_name)                         'trimming the client name
		IF instr(client_name, ",") <> 0 THEN client_name =  replace(client_name, ",", ", ")
		objExcel.Cells(excel_row, 3).Value = trim(client_name)

        EMReadScreen client_ssn, 11, 5, 13						'Reads the client name and adds to excel
	    objExcel.Cells(excel_row, 5).Value = trim(client_ssn)

		EMReadScreen earner_name_match, 6, 10, 3
		IF earner_name_match = "EARNER" THEN
			EMReadScreen earner_name, 35, 10, 16
			'Formatting the client name for the spreadsheet
			earner_name = trim(earner_name)                         'trimming the client name
			DO
				IF instr(earner_name, "  ") <> 0 THEN earner_name = replace(earner_name, "  ", " ")
			LOOP UNTIL instr(earner_name, "  ") = 0
			objExcel.Cells(excel_row, 4).Value = trim(earner_name)
		END IF
        EMReadScreen active_programs, 5, 7, 13			'Reads the client name and adds to excel
	    active_programs = trim(active_programs)
	    objExcel.Cells(excel_row, 12).Value = active_programs

	    programs = ""
	    	IF instr(active_Programs, "D") THEN programs = programs & "DWP, "
	    	IF instr(active_Programs, "F") THEN programs = programs & "Food Support, "
	    	IF instr(active_Programs, "H") THEN programs = programs & "Health Care, "
	    	IF instr(active_Programs, "M") THEN programs = programs & "Medical Assistance, "
	    	IF instr(active_Programs, "S") THEN programs = programs & "MFIP, "
	    programs = trim(programs)'trims excess spaces of programs
	    IF right(programs, 1) = "," THEN programs = left(programs, len(programs) - 1)'takes the last comma off of programs when autofilled into dialog

	    EMSearch "SEND IEVS DIFFERENCE NOTICE?", row, col 	'Finds where the difference notice code is - because it moves
	    EMReadScreen diff_notc_sent, 1, row, 36				'Reads if diff notice was sent or not
	    If diff_notc_sent = "N" Then diff_notc_date = ""
	    If diff_notc_sent = "Y" Then EMReadScreen diff_notc_date, 8, row, 72	'If notice was sent, reads the date it was sent
	    objExcel.Cells(excel_row, 13).Value = diff_notc_sent	'Adding both of these to excel
	    objExcel.Cells(excel_row, 14).Value = diff_notc_date

	    IF match_type = "A30" or match_type = "A40" THEN 'SDXS & BNDX
	    	EMReadScreen income_amount, 15, 9, 18			'Reads the income amount and adds to excel
	    	income_amount = trim(income_amount)
	    	If instr(income_amount, "NOT") THEN 					  'establishing the length of the variable
	    		position = InStr(income_amount, "NOT")    		      'sets the position at the deliminator
	    		income_amount = left(income_amount, position - 1)  'establishes employer as being before the deliminator
	    	END IF
	    	income_amount = replace(income_amount, "$", "")
	    	objExcel.Cells(excel_row, 15).Value = income_amount
	    END IF

	    IF match_type = "A50" or match_type = "A51" THEN 'WAGE'
	    	EMReadScreen match_year, 4, 9, 16			'Reads the match_year and adds to excel
	    	match_year = trim(match_year)
	    	objExcel.Cells(excel_row, 16).Value = match_year
	    	EMReadScreen income_source, 60, 9, 31			'Reads the income_source and adds to excel
	    	income_source = trim(income_source)							  'establishing the length of the variable
	    	length = len(income_source)
	    	position = InStr(income_source, "AMT: $")    		      'sets the position at the deliminator
	    	income_source = left(income_source, position - 1 )  'establishes employer as being before the deliminator
	    	objExcel.Cells(excel_row, 17).Value = income_source
	    	EMSearch "AMT: $", 9, col
	    	'MsgBox col
	    	EMReadScreen income_amount, 72 - col, 9, col + 6			'Reads the income_amount and adds to excel up to 36 spaces
	    	'MsgBox 81 - col & vbcr & income_amount
	    	income_amount = trim(income_amount)
	    	objExcel.Cells(excel_row, 15).Value = income_amount
	    END IF

	    IF match_type = "A60" THEN 'UBEN'
	    	EMReadScreen nonwage_date, 10, 9, 39			'Reads the nonwage_date and adds to excel
	    	nonwage_date = trim(nonwage_date)
	    	objExcel.Cells(excel_row, 18).Value = nonwage_date
	    	EMReadScreen income_amount, 20, 9, 11			'Reads the income_amount and adds to excel
	    	income_amount = trim(income_amount)
	    	If instr(income_amount, "DATE") THEN 					  'establishing the length of the variable
	    		position = InStr(income_amount, "DATE")    		      'sets the position at the deliminator
	    		income_amount = left(income_amount, position - 1)  'establishes income_amount as being before the deliminator
	    	END IF
	    	income_amount = replace(income_amount, "$", "")
	    	objExcel.Cells(excel_row, 15).Value = income_amount
	    END IF

	    IF match_type = "A70" THEN 'BEER'
	    	EMReadScreen match_year, 2, 9, 9			'Reads the match_year and adds to excel
	    	match_year = trim(match_year)
	    	objExcel.Cells(excel_row, 16).Value = match_year
	    	EMReadScreen income_source, 60, 9, 22			'Reads the income_source and adds to excel
	    	income_source = trim(income_source)
	    	If instr(income_source, "AMOUNT: $") THEN 					  'establishing the length of the variable
	    	    position = InStr(income_source, "AMOUNT: $")    		      'sets the position at the deliminator
	    	    income_source = left(income_source, position - 1)  'establishes income_source as being before the deliminator
	    	END IF
	    	objExcel.Cells(excel_row, 17).Value = income_source
	    	EMSearch "AMOUNT: $", 9, col
	    	EMReadScreen income_amount, 20, 9, col + 9			'Reads the income_amount and adds to excel
	    	income_amount = trim(income_amount)
	    	If instr(income_amount, "AMOUNT: $") THEN 					  'establishing the length of the variable
	    	    position = InStr(income_amount, "AMOUNT: $")    		      'sets the position at the deliminator
	    	    income_amount = right(income_amount, position)  'establishes income_amount as being before the deliminator
	    	END IF
	    	objExcel.Cells(excel_row, 15).Value = income_amount
	    END IF

	    IF match_type = "A80" THEN 'UNVI '
	    	EMReadScreen match_year, 4, 9, 9			'Reads the match_year and adds to excel
	    	match_year = trim(match_year)
	    	objExcel.Cells(excel_row, 16).Value = match_year
	    	EMReadScreen income_amount, 20, 9, 33			'Reads the income_amount and adds to excel
	    	income_amount = trim(income_amount)
	    	income_amount = replace(income_amount, "$", "")
	    	objExcel.Cells(excel_row, 15).Value = income_amount
	    END IF

	    PF3 'back to the list'
	   IEVC_Row = IEVC_Row + 1 'increment to the next row on the panel
    End if
	If IEVC_Row = 18 Then 		'If we have reached the end of the page, it will go to the next page
		PF8
		IEVC_Row = 8			'Resets the row
		EMReadScreen last_page_check, 21, 24, 2
	End If
	excel_row = excel_row + 1	'increments the excel row so we don't overwrite our data
	STATS_counter = STATS_counter + 1		'Counts 1 item for every Match found and entered into excel.			diff_notc_date = ""			'blanks this out so that the information is not carried over in the do-loop'
	maxis_case_number = ""
Loop until last_page_check = "THIS IS THE LAST PAGE"

STATS_counter = STATS_counter - 1 'removed increment from the start of the script for an accurate count

'Centers the text for the columns with days remaining and difference notice
objExcel.Columns(6).HorizontalAlignment = -4108
objExcel.Columns(7).HorizontalAlignment = -4108
objExcel.Columns(8).HorizontalAlignment = -4108

excel_is_not_blank = chr(34) & "<>" & chr(34)		'Setting up a variable for useable quote marks in Excel

'Query date/time/runtime info
objExcel.Cells(2, 22).Font.Bold = TRUE
objExcel.Cells(3, 22).Font.Bold = TRUE
objExcel.Cells(4, 22).Font.Bold = TRUE
objExcel.Cells(5, 22).Font.Bold = TRUE

ObjExcel.Cells(2, 22).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, 23).Value = now
ObjExcel.Cells(3, 22).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(3, 23).Value = timer - query_start_time
ObjExcel.Cells(4, 22).Value = "Number of IEVS with No DAYS remaining:"
objExcel.Cells(4, 23).Value = "=COUNTIFS(I:I, " & Chr(34) & "<=0" & Chr(34) & ", I:I, " & excel_is_not_blank & ")"	'Excel formula
ObjExcel.Cells(5, 22).Value = "Number of total UNRESOLVED IEVS:"
objExcel.Cells(5, 23).Value = "=(COUNTIF(I:I, " & excel_is_not_blank & ")-1)"	'Excel formula

'Formatting the column width.
FOR i = 1 to 23
	objExcel.Columns(i).AutoFit()
NEXT

script_end_procedure("Success! The spreadsheet has all requested information.")
