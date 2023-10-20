'Required for statistical purposes==========================================================================================
name_of_script = "Test.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 500                     'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
'END OF stats block================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
    IF on_the_desert_island = TRUE Then
        FuncLib_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\desert-island\MASTER FUNCTIONS LIBRARY.vbs"
        Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
        Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
        text_from_the_other_script = fso_command.ReadAll
        fso_command.Close
        Execute text_from_the_other_script
    ELSEIF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
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

    '---------------------------

function new_create_mainframe_friendly_date(date_variable, screen_row, screen_col, variable_length)
'--- This function creates a mainframe friendly date. This can be used for both year formats and input spacing.
'~~~~~ date_variable: the name of the variable to output
'~~~~~ screen_row: row to start writing date
'~~~~~ screen_col: column to start writing date
'~~~~~ variable length: the number of days to add to the date_variable for output. Entering "tikl_date" will calculate the first date for negative action allowing for 10 day notice. 
'      entering any other non-numeric variable will result in 0 days added.
'===== Keywords: MAXIS, PRISM, MMIS, date, create
'Year type is now variable length. This is a date offset calculation in days.
	IF isnumeric(variable_length) = false THEN variable_length = 0 'Need to handle the old function, where this value was "YY or YYYY" for handling year.
	date_variable = dateadd("d", variable_length, date_variable) 'adding the number of days to the date

	'Formatting the parts of the Date to correct lengths
	var_month = datepart("m", date_variable)
	IF len(var_month) = 1 THEN var_month = "0" & var_month
	var_day = datepart("d", date_variable)
	IF len(var_day) = 1 THEN var_day = "0" & var_day
	var_year = datepart("yyyy", date_variable)

	'This section reads the location we're entering the date and determines format
	EMREadScreen date_space, 14, screen_row, screen_col 
	EMReadScreen MMIS_check, 5, 1, 2
	'This section determines what fields we're dealing with
	EMSetCursor screen_row, screen_col
	EMSendKey "<TAB>"
	EMGetCursor row_two, column_two
	EMSendKey "<TAB>"
	EMGetCursor row_three, column_three
	IF (row_two = screen_row AND column_two - screen_col > 6) OR screen_row <> row_two THEN 'This is a single field for the date
		EMWriteScreen "_", screen_row, screen_col +8 'We're determining if this is a writeable field for year format. Have to write a char because MMIS doesn't mark fields
		EMREadScreen year_field, 1, screen_row, screen_col +8
		If year_field = "_" Then 'This is a 4 digit year space
			EMWriteScreen var_month & "/" & var_day & "/" & var_year, screen_row, screen_col
		Else '2 digit year
			EMWriteScreen var_month & "/" & var_day & "/" & right(var_year, 2), screen_row, screen_col
		End If
	ElseIf row_three = screen_row AND column_two - screen_col = 5 AND column_three - column_two = 5 THEN 'This is 3 spaces between 3 date parts
		EMWriteScreen var_month, screen_row, screen_col
		EMWriteScreen var_day, screen_row, screen_col + 5
		EMREadScreen year_field, 1, screen_row + 12
		IF year_field = " " Then
			EMWriteScreen right(var_year, 2), screen_row, column_three
		Else
			EMWriteScreen var_year, screen_row, column_three
		End If
	ElseIf (row_three = screen_row AND column_two - screen_col = 3 AND column_three - column_two = 3) OR (row_two = screen_row AND column_two - screen_col = 6) THEN 'This is 1 space between 3 date parts or center date is an "01" in system
		EMWriteScreen var_month, screen_row, screen_col
		EMWriteScreen var_day, screen_row, screen_col + 3
		EMREadScreen year_field, 1, screen_row, screen_col + 8  'check for year format
		If year_field = " " Then
			EMWriteScreen right(var_year, 2), screen_row, screen_col + 6
		Else
			EMWriteScreen var_year, screen_row, screen_col + 6
		End If
	ElseIf column_two - screen_col = 3 AND (column_three - column_two > 3 OR row_three <> row_two) Then'Month, space, year
		EMWriteScreen var_month, screen_row, screen_col
		EMREadScreen year_field, 1, screen_row, screen_col + 6  'check for year format
		If year_field = " " Then
			EMWriteScreen right(var_year, 2), screen_row, screen_col + 3
		Else
			EMWriteScreen var_year, screen_row, screen_col + 3
		End If
	ElseIf column_two - screen_col = 5 AND (column_three - column_two > 5 OR row_three <> row_two) Then'3 spaces between month/year (TRAC)
		EMWriteScreen var_month, screen_row, screen_col
		EMREadScreen year_field, 1, screen_row, screen_col + 7  'check for year format
		If year_field = " " Then
			EMWriteScreen right(var_year, 2), screen_row, screen_col + 5
		Else
			EMWriteScreen var_year, screen_row, screen_col + 5
		End If
	Else
		MsgBox "Something went wrong. The script has encountered an unsupported date field or format."
	End If 

end function

EMConnect ""

call new_create_mainframe_friendly_date(date, 8, 27, 365)
call new_create_mainframe_friendly_date("10/01/23", 9, 71, 0)


stopscript