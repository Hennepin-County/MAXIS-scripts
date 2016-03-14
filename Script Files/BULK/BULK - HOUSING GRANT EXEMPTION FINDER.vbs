'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - HOUSING GRANT EXEMPTION FINDER.vbs"
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
STATS_manualtime = 120                     'manual run time in seconds
STATS_denomination = "C"       			   'C is for each CASE
'END OF stats block==============================================================================================

'DIALOGS----------------------------------------------------------------------
BeginDialog pull_REPT_data_into_excel_dialog, 0, 0, 218, 120, "Housing Grant Exemption Finder"
  EditBox 84, 20, 130, 15, worker_number
  CheckBox 4, 65, 150, 10, "Check here to run this query county-wide.", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 109, 100, 50, 15
    CancelButton 164, 100, 50, 15
  Text 4, 25, 65, 10, "Worker(s) to check:"
  Text 4, 80, 210, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
  Text 14, 5, 125, 10, "***PULL REPT DATA INTO EXCEL***"
  Text 4, 40, 210, 20, "Enter all 7 digits of your workers' x1 numbers (ex: x######), separated by a comma."
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""

'remove after testing
worker_number = "x127EZ4"

'Shows dialog
DO
	Dialog pull_rept_data_into_Excel_dialog
	If buttonpressed = cancel then stopscript
	If worker_number = "" then MsgBox "Please enter at least one worker number." 
LOOP until worker_number <> ""

'Starting the query start time (for the query runtime at the end)
query_start_time = timer
Call check_for_MAXIS(True)	'Checking for MAXIS

'Fun with dates! --Creating variables for the rolling 12 calendar months
'current month -1
CM_minus_1_mo =  right("0" &          	 DatePart("m",           DateAdd("m", -1, date)            ), 2)
CM_minus_1_yr =  right(                  DatePart("yyyy",        DateAdd("m", -1, date)            ), 2)
'current month -2'
CM_minus_2_mo =  right("0" &             DatePart("m",           DateAdd("m", -2, date)            ), 2)
CM_minus_2_yr =  right(                  DatePart("yyyy",        DateAdd("m", -2, date)            ), 2)
'current month -3'
CM_minus_3_mo =  right("0" &             DatePart("m",           DateAdd("m", -3, date)            ), 2)
CM_minus_3_yr =  right(                  DatePart("yyyy",        DateAdd("m", -3, date)            ), 2)
'current month -4'
CM_minus_4_mo =  right("0" &             DatePart("m",           DateAdd("m", -4, date)            ), 2)
CM_minus_4_yr =  right(                  DatePart("yyyy",        DateAdd("m", -4, date)            ), 2)
'current month -5'
CM_minus_5_mo =  right("0" &             DatePart("m",           DateAdd("m", -5, date)            ), 2)
CM_minus_5_yr =  right(                  DatePart("yyyy",        DateAdd("m", -5, date)            ), 2)
'current month -6'
CM_minus_6_mo =  right("0" &             DatePart("m",           DateAdd("m", -6, date)            ), 2)
CM_minus_6_yr =  right(                  DatePart("yyyy",        DateAdd("m", -6, date)            ), 2)
'current month -7'
CM_minus_7_mo =  right("0" &             DatePart("m",           DateAdd("m", -7, date)            ), 2)
CM_minus_7_yr =  right(                  DatePart("yyyy",        DateAdd("m", -7, date)            ), 2)
'current month -8'
CM_minus_8_mo =  right("0" &             DatePart("m",           DateAdd("m", -8, date)            ), 2)
CM_minus_8_yr =  right(                  DatePart("yyyy",        DateAdd("m", -8, date)            ), 2)
'current month -9'
CM_minus_9_mo =  right("0" &             DatePart("m",           DateAdd("m", -9, date)            ), 2)
CM_minus_9_yr =  right(                  DatePart("yyyy",        DateAdd("m", -9, date)            ), 2)
'current month -10'
CM_minus_10_mo =  right("0" &            DatePart("m",           DateAdd("m", -10, date)           ), 2)
CM_minus_10_yr =  right(                 DatePart("yyyy",        DateAdd("m", -10, date)           ), 2)
'current month -11'
CM_minus_11_mo =  right("0" &            DatePart("m",           DateAdd("m", -11, date)           ), 2)
CM_minus_11_yr =  right(                 DatePart("yyyy",        DateAdd("m", -11, date)           ), 2)

'Establishing value of variables
current_month = CM_mo & "/" & CM_yr
current_month_minus_one = CM_minus_1_mo & "/" & CM_minus_1_yr
current_month_minus_two = CM_minus_2_mo & "/" & CM_minus_2_yr
current_month_minus_three = CM_minus_3_mo & "/" & CM_minus_3_yr
current_month_minus_four = CM_minus_4_mo & "/" & CM_minus_4_yr
current_month_minus_five = CM_minus_5_mo & "/" & CM_minus_5_yr
current_month_minus_six = CM_minus_6_mo & "/" & CM_minus_6_yr
current_month_minus_seven = CM_minus_7_mo & "/" & CM_minus_7_yr
current_month_minus_eight = CM_minus_8_mo & "/" & CM_minus_8_yr
current_month_minus_nine = CM_minus_9_mo & "/" & CM_minus_9_yr
current_month_minus_ten = CM_minus_10_mo & "/" & CM_minus_10_yr
current_month_minus_eleven = CM_minus_11_mo & "/" & CM_minus_11_yr

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Setting the Excel rows with varibles 
ObjExcel.Cells(1, 1).Value = "WORKER"
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
ObjExcel.Cells(1, 3).Value = "NAME"
ObjExcel.Cells(1, 4).Value = "EMPS"
ObjExcel.Cells(1, 5).Value = "DISA DATES"
ObjExcel.Cells(1, 6).Value = "MFIP BEGIN DATE"
ObjExcel.Cells(1, 7).Value = current_month					'using date calculations above, list will generate a rolling 12 months of information
ObjExcel.Cells(1, 8).Value = current_month_minus_one
ObjExcel.Cells(1, 9).Value = current_month_minus_two
ObjExcel.Cells(1, 10).Value = current_month_minus_three
ObjExcel.Cells(1, 11).Value = current_month_minus_four
ObjExcel.Cells(1, 12).Value = current_month_minus_five
ObjExcel.Cells(1, 13).Value = current_month_minus_six
ObjExcel.Cells(1, 14).Value = current_month_minus_seven
ObjExcel.Cells(1, 15).Value = current_month_minus_eight
ObjExcel.Cells(1, 16).Value = current_month_minus_nine
ObjExcel.Cells(1, 17).Value = current_month_minus_ten
ObjExcel.Cells(1, 18).Value = current_month_minus_eleven

FOR i = 1 to 18		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	objExcel.Columns(i).AutoFit()						'sizing the colums'
NEXT

'Figuring out what to put in each Excel col. To add future variables to this, add the checkbox variables below and copy/paste the same code!
'Below, use the "[blank]_col" variable to recall which col you set for which option.
col_to_use = 19 'Starting with 19 because cols 1-18 are already used

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

'Setting the variable for what's to come
excel_row = 2

For each worker in worker_array
	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
	EMWriteScreen CM_mo, 20, 43
	EMWriteScreen CM_yr, 20, 46
	Call navigate_to_MAXIS_screen("REPT", "MFCM")
	EMWriteScreen worker, 21, 13
	transmit
	'Skips workers with no info
	EMReadScreen has_content_check, 29, 7, 6
    has_content_check = trim(has_content_check)
	If has_content_check <> "" then
		'Grabbing each case number on screen
		Do
			MAXIS_row = 7	'Set variable for next do...loop
			Do
				EMReadScreen emps_status, 2, MAXIS_row, 52		'Reading Emps Status
				If emps_status = "02" OR emps_status = "08" OR _	
				   	emps_status = "10" OR emps_status = "12" OR _
				   	emps_status = "23" OR emps_status = "24" OR _
				   	emps_status = "27" OR emps_status = "15" OR _
				   	emps_status = "18" OR emps_status = "18" OR _
				   	emps_status = "30" OR emps_status = "33" THEN
						EMReadScreen case_number, 8, MAXIS_row, 6		  'Reading case number
						EMReadScreen client_name, 20, MAXIS_row, 16		'Reading client name
						'if more than one HH member is on the list then non-MEMB 01's don't have a case number listed, this fixes that
						If trim(case_number) = "" AND trim(client_name) <> "" then 			'if there's a name and no case number 
							EMReadScreen alt_case_number, 8, MAXIS_row - 1, 6				'then it reads the row above
							case_number = alt_case_number									'restablishes that in this instance, alt case number = case number'
						END IF
						'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
						If trim(case_number) <> "" and instr(all_case_numbers_array, case_number) <> 0 then exit do
						all_case_numbers_array = trim(all_case_numbers_array & " " & case_number)
						If trim(case_number) = "" and trim(client_name) = "" then exit do			'Exits do if we reach the end					
					
						'add case/case information to Excel
        			ObjExcel.Cells(excel_row, 1).Value = worker
        			ObjExcel.Cells(excel_row, 2).Value = case_number
        			ObjExcel.Cells(excel_row, 3).Value = client_name
        			ObjExcel.Cells(excel_row, 4).Value = emps_status
    				excel_row = excel_row + 1	'moving excel row to next row'
					STATS_counter = STATS_counter + 1		 'adds one instance to the stats counter
					case_number = ""
				END IF
							'Blanking out variable
				MAXIS_row = MAXIS_row + 1	'adding one row to search for in MAXIS		
			Loop until MAXIS_row = 19
			PF8
			EMReadScreen last_page_check, 21, 24, 2	'Checking for the last page of cases.
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if                     
next

'Now the script checks the potentially exempt members for subsidized housing
excel_row = 2           're-establishing the row to start checking the members for 
Do 						'Loops until there are no more cases in the Excel list
	case_number = objExcel.cells(excel_row, 2).Value	're-establishing the case numbers
	client_name = objExcel.cells(excel_row, 3).Value	're-establishing the client name
	client_name = trim(client_name)
	If case_number = "" then exit do	
	Call navigate_to_MAXIS_screen("STAT", "MEMB")
	DO					'Do Loop is searching for MEMB #, this is needed for SHEL and DISA
        EMReadScreen last_name, 25, 6, 30
        EMReadScreen first_name, 12, 6, 63
        EmReadScreen middle_name, 1, 6, 79
		last_name = Replace(last_name,"_","")
        first_name = Replace(first_name,"_","")
		middle_name= Replace(middle_name,"_","")
		name_of_client = trim(last_name & ", " & first_name & " " & middle_name) 		'adds the variables together to make a new name variable
		If ucase(client_name) = ucase(name_of_client) then 
			EMReadScreen memb_number, 2, 4, 33
			exit do
		ELSE 
			transmit
		END IF
		Msgbox client_name & vbnewline & name_of_client
    LOOP until client_name = name_of_client		'if not, transmit and loop until client_name is found
    
    'checking member for subsidized housing 
    Call navigate_to_MAXIS_screen("STAT", "SHEL")
    EMWriteScreen memb_number, 20, 76
    transmit
    EmReadScreen sub_housing, 1, 6, 46
    If sub_housing <> "Y" then 				'if member doesn't have sub housing, then the sub housing fiater is not necessary
        'Deleting the blank results to clean up the spreadsheet
        SET objRange = objExcel.Cells(excel_row, 1).EntireRow
        objRange.Delete
        excel_row = excel_row - 1	
    END IF
    excel_row = excel_row + 1
LOOP UNTIL objExcel.Cells(excel_row, 1).Value = ""	'looping until the list is complete

'Now the script checks for MFIP start date, disa dates and inputs payment history
'STAT PROG PORTION
excel_row = 2           're-establishing the row to start checking the members for 
DO	
	case_number = objExcel.cells(excel_row, 2).Value	're-establishing the case numbers
	Call navigate_to_MAXIS_screen("STAT", "PROG")
	'reading the MFIP start date
	EMReadScreen prog_one, 2, 6, 67				'checking 1st line of CASH PROG for elig MFIP
	EMReadScreen prog_status_one, 4, 6, 74
	EMReadScreen elig_begin_date_one, 8, 6, 44
	If prog_one = "MF" and prog_status_one = "ACTV" then 
		elig_begin_date_one = Replace(elig_begin_date_one," ","/")
		elig_begin_date = elig_begin_date_one
	ELSE
		EMReadScreen prog_two, 2, 7, 67				'checking 2nd line of CASH PROG for elig MFIP if 1st line isn't MFIP
		EMReadScreen prog_status_two, 4, 7, 74
		EMReadScreen elig_begin_date_two, 8, 7, 44
		elig_begin_date = elig_begin_date_two
	END IF 
	ObjExcel.Cells(excel_row, 6).Value = elig_begin_date
	msgbox elig_begin_date
	
	'STAT DISA PORTION
	Call navigate_to_MAXIS_screen("STAT", "DISA")
	EMWriteScreen memb_number, 20, 76
	transmit
	'Reading the disa dates
	EMReadScreen disa_start_date, 10, 6, 47
	EMReadScreen disa_end_date, 10, 6, 69
	disa_start_date = Replace(disa_start_date," ","/")
	disa_end_date = Replace(disa_end_date," ","/")
	disa_dates = trim(disa_start_date) & " - " & trim(disa_end_date)
	'adding disa date to Excel 
	ObjExcel.Cells(excel_row, 5).Value = disa_dates
	
	'INQX/INQD PORTION
	back_to_self
	EMWriteScreen "________", 18, 43
	EMWriteScreen case_number, 18, 43
	Call navigate_to_MAXIS_screen("MONY", "INQX")
	EMWriteScreen CM_minus_11_mo, 6, 38		'entering footer month/year 11 months prior to see full year
	EMWriteScreen CM_minus_11_yr, 6, 41
	EMWriteScreen CM_mo, 6, 53		'entering current footer month/year
	EMWriteScreen CM_yr, 6, 56
	EMWriteScreen "x", 10, 5		'selecting MFIP
	transmit
	'searching for the housing grant issued on the INQD screen(s)
	DO
		row = 6
		DO
			EMReadScreen housing_grant, 2, row, 19		'searching for housing grant issuance
			If housing_grant = "__" then exit do
			IF housing_grant = "HG" then
				'reading the housing grant information
				EMReadScreen EMER_type, 2, row, 19
				EMReadScreen EMER_amt_issued, 7, row, 39
				EMReadScreen HG_month, 2, row, 73
				EMReadScreen HG_year, 2, row, 77
			END IF	
			row = row + 1
		Loop until row = 18				'repeats until the end of the page
			PF8
			EMReadScreen last_page_check, 21, 24, 2
			If last_page_check <> "THIS IS THE LAST PAGE" then row = 6		're-establishes row for the new page
	LOOP UNTIL last_page_check = "THIS IS THE LAST PAGE"
		
	'excel_row = excel_row + 1	'moving excel row to next row'
'LOOP UNTIL objExcel.Cells(excel_row, 1).Value = ""	'looping until the list is complete
	
'------------------------------Post MAXIS coding-----------------------------------------------------------------------------
col_to_use = col_to_use + 2	'Doing two because the wrap-up is two columns

'Query date/time/runtime info
objExcel.Cells(1, col_to_use - 1).Font.Bold = TRUE
objExcel.Cells(2, col_to_use - 1).Font.Bold = TRUE
ObjExcel.Cells(1, col_to_use - 1).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, col_to_use).Value = now
ObjExcel.Cells(2, col_to_use - 1).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, col_to_use).Value = timer - query_start_time

'Autofitting columns
For col_to_autofit = 1 to col_to_use
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

'Logging usage stats
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
msgbox STATS_counter
script_end_procedure("Success! Please review the list generated.")