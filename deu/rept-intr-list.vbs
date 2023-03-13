'Required for statistical purposes===============================================================================
name_of_script = "BULK - REPT-INTR LIST.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 35                               'manual run time, per line, in seconds
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
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

EMConnect "" 'Connects to MAXIS

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 176, 75, "REPT-INTR List"
  ButtonGroup ButtonPressed
    OkButton 65, 55, 50, 15
    CancelButton 120, 55, 50, 15
  GroupBox 10, 5, 160, 45, "About this script:"
  Text 15, 20, 150, 25, "This script gathers a list of all the PARIS matches in the agency. Press OK to continue or CANCEL to close the script. "
EndDialog

'Shows dialog
Do
  	dialog Dialog1
  	cancel_without_confirmation
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call check_for_MAXIS(False) 'Checking for MAXIS

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Name for the current sheet'
ObjExcel.ActiveSheet.Name = "REPT-INTR"

'Excel headers and formatting the columns
objExcel.Cells(1, 1).Value     = "CASE NUMBER"
objExcel.Cells(1, 1).Font.Bold = TRUE
objExcel.Cells(1, 2).Value     = "WORKER NUMBER"
objExcel.Cells(1, 2).Font.Bold = TRUE
objExcel.Cells(1, 3).Value     = "PMI"
objExcel.Cells(1, 3).Font.Bold = TRUE
objExcel.Cells(1, 4).Value     = "APPLICANT NAME"
objExcel.Cells(1, 4).Font.Bold = TRUE
objExcel.Cells(1, 5).Value     = "MONTH"
objExcel.Cells(1, 5).Font.Bold = TRUE
objExcel.Cells(1, 6).Value     = "RESOLUTION"
objExcel.Cells(1, 6).Font.Bold = TRUE
objExcel.Cells(1, 7).Value     = "NOTICE DATE"
objExcel.Cells(1, 7).Font.Bold = TRUE

'Sets variable for all of the Excel stuff
excel_row = 2
back_to_SELF
CALL navigate_to_MAXIS_screen("REPT", "INTR")		'Go to REPT INTR
CALL clear_line_of_text(5, 15)
CALL clear_line_of_text(6, 15)
EMWriteScreen "AL", 5, 67			'Entering the resolution code selected in dialog
TRANSMIT

If start_month <> "" Then
	start_month = right("00" & start_month, 2)
	EMWriteScreen start_month, 6, 67			'Entering the date range if selected
End If
If start_year <> "" Then
	start_year = right("00" & start_year, 2)
	EMWriteScreen start_year, 6, 70
End If
If end_month <> "" Then
	end_month = right("00" & end_month, 2)
	EMWriteScreen end_month, 7, 67
End If
If end_year <> "" Then
	end_year = right("00" & end_year, 2)
	EMWriteScreen end_year, 7, 70
End If
transmit

EMReadScreen intr_exists, 8, 11, 5				'Looking if there are any matches listed under this worker
intr_exists = trim(intr_exists)
row = 11
If intr_exists <> "" Then 	'If there are any matches the script will pull detail
	Do
		EMReadScreen maxis_case_number, 8, row, 5			'Reading the case number
		maxis_case_number = trim(maxis_case_number)			'removing the spaces
		If maxis_case_number = "" then exit Do 		'Once the script reaches the last line in the list, it will go to the next worker

		EMReadScreen worker_number, 7, row, 14				'Reading the worker x-number listed on the match - necessary if the number in the array is a supervisor number
		EMReadScreen PMI_number, 7, row, 23
		EMReadScreen client_name, 20, row, 31				'Reading the client name and removing the blanks
		client_name = trim(client_name)
		client_name = UCASE(client_name)
		EMReadScreen match_month, 5, row, 53				'Reading the month the match was issued
		match_month = replace(match_month, " ", "/01/")		'Formatting the date as a date for entry into Excel
		EMReadScreen res_status, 2, row, 64					'Reading the resolution status
		EMReadScreen notice_date, 8, row, 71				'Reading the notice date field
		if notice_date = "        " then notice_date = ""	'blanking out if there is no date
		notice_date = replace(notice_date, " ", "/")		'Formatting the date

		'Adding all the information to Excel
		objExcel.Cells(excel_row, 1).Value = maxis_case_number
		objExcel.Cells(excel_row, 2).Value = worker_number
		objExcel.Cells(excel_row, 3).Value = PMI_number
		objExcel.Cells(excel_row, 4).Value = client_name
		objExcel.Cells(excel_row, 5).Value = match_month
		objExcel.Cells(excel_row, 6).Value = res_status
		objExcel.Cells(excel_row, 7).Value = notice_date

		row = row + 1		'Go to the next excel row
		If row = 19 Then 		'If we have reached the end of the page, it will go to the next page
			PF8
			row = 11			'Resets the row
			EMReadScreen last_page_check, 21, 24, 2
		End If
		excel_row = excel_row + 1	'increments the excel row so we don't overwrite our data
		STATS_counter = STATS_counter + 1		'Counts 1 item for every Match found and entered into excel.			diff_notc_date = ""			'blanks this out so that the information is not carried over in the do-loop'
		maxis_case_number = ""
	Loop until last_page_check = "THIS IS THE LAST PAGE"
End If

'Centers the text for the columns with days remaining and difference notice
objExcel.Columns(6).HorizontalAlignment = -4108
objExcel.Columns(7).HorizontalAlignment = -4108
objExcel.Columns(8).HorizontalAlignment = -4108

'excel_is_not_blank = chr(34) & "<>" & chr(34)		'Setting up a variable for useable quote marks in Excel

For col_to_autofit = 1 to 7
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

STATS_counter = STATS_counter - 1		'removing the initial counter so that this number is correct.
script_end_procedure("Success! The spreadsheet has all requested information.")
