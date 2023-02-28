'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - FSS INFO.vbs"
start_time = timer
STATS_counter = 1                       'sets the stats counter at one
STATS_manualtime = "100"                'manual run time in seconds
STATS_denomination = "C"       			'C is for each CASE
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
call changelog_update("05/19/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 218, 120, "FSS information dialog"
  EditBox 75, 20, 135, 15, worker_number
  CheckBox 5, 65, 150, 10, "Check here to run this query county-wide.", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 105, 105, 50, 15
    CancelButton 160, 105, 50, 15
  Text 5, 25, 65, 10, "Worker(s) to check:"
  Text 5, 80, 210, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
  Text 10, 5, 125, 10, "***PULL REPT DATA INTO EXCEL***"
  Text 5, 40, 210, 20, "Enter all 7 digits of your workers' x1 numbers (ex: x######), separated by a comma."
EndDialog
'Shows dialog
Do
	Do
		Dialog Dialog1
		Cancel_without_confirmation
		If (all_workers_check = 0 AND worker_number = "") then MsgBox "Please enter at least one worker number." 'allows user to select the all workers check, and not have worker number be ""
	LOOP until all_workers_check = 1 or worker_number <> ""
	Call check_for_password(are_we_passworded_out)
Loop until check_for_password(are_we_passworded_out) = False		'loops until user is password-ed out

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Setting the Excel rows with variables
ObjExcel.Cells(1, 1).Value = "WORKER"
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
ObjExcel.Cells(1, 3).Value = "EMPS"
ObjExcel.Cells(1, 4).Value = "CLIENT NAME"
ObjExcel.Cells(1, 5).Value = "REF #"
ObjExcel.Cells(1, 6).Value = "DISA DATES"
objExcel.cells(1, 7).Value = "Privileged Cases"

ObjExcel.columns(5).NumberFormat = "00" 'formatting the worksheet to show 2 digit member numbers when the 1st number is 0

FOR i = 1 to 7		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

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

'establishing the row to start searching in the Excel spreadsheet
excel_row = 2

For each worker in worker_array
	back_to_self
	EMWriteScreen CM_mo, 20, 43				'
	EMWriteScreen CM_yr, 20, 46
	Call navigate_to_MAXIS_screen("REPT", "MFCM")			'navigates to MFCM in the current footer month/year'
	EMWriteScreen worker, 21, 13
	transmit

	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason----'Skips workers with no info
	EMReadScreen has_content_check, 29, 7, 6
    has_content_check = trim(has_content_check)
	If has_content_check <> "" then
		Do
			MAXIS_row = 7	'Sets the row to start searching in MAXIS for
			Do
				EMReadScreen emps_status, 2, MAXIS_row, 52		'Reading Emps Status & only searches for exempt emps status codes
				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 6  	'Reading case number
				EMReadScreen client_name, 18, MAXIS_row, 16
                
                'if more than one HH member is on the list then non-MEMB 01's don't have a case number listed, this fixes that
				If trim(MAXIS_case_number) = "" AND trim(emps_status) <> "" then 			'if there's a name and no case number
					EMReadScreen alt_case_number, 8, MAXIS_row - 1, 6				'then it reads the row above
                    MAXIS_case_number = alt_case_number									'restablishes that in this instance, alt case number = case number'    
                END IF
                
                If trim(MAXIS_case_number) = "" and trim(emps_status) = "" then exit do			'Exits do if we reach the end
				
				'add case/case information to Excel
        		ObjExcel.Cells(excel_row, 1).Value = worker
        		ObjExcel.Cells(excel_row, 2).Value = trim(MAXIS_case_number)
    			ObjExcel.Cells(excel_row, 3).Value = emps_status
                ObjExcel.Cells(excel_row, 4).Value = trim(client_name)		'adds client name to Excel list
                
				excel_row = excel_row + 1	'moving excel row to next row'
				MAXIS_case_number = ""          'Blanking out variable
				MAXIS_row = MAXIS_row + 1	'adding one row to search for in MAXIS
			Loop until MAXIS_row = 19
			PF8
			EMReadScreen last_page_check, 21, 24, 2	'Checking for the last page of cases.
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
next

'Now the script goes back into MFCM and grabs the member # and client name, then cchecks the potentially exempt members for subsidized housing
excel_row = 2           're-establishing the row to start checking the members for
Do
	MAXIS_case_number  = objExcel.cells(excel_row, 2).Value	're-establishing the case number to use for the case
    client_name        = objExcel.cells(excel_row, 4).Value	're-establishing the client name to use for the case
    If MAXIS_case_number = "" then exit do						'exits do if the case number is ""
	Call navigate_to_MAXIS_screen("REPT", "MFCM")
	
    EMReadScreen case_content, 7, 8, 7
	If trim(case_content) = "" then 
		'making sure we are getting the right person for cases where there are more than one case. 
    	row = 7
    	Do 
        	EMReadScreen case_name, 18, row, 16
			case_name = trim(case_name)
        	If case_name <> client_name then row = row + 1
    	LOOP until case_name = client_name  
		EMWriteScreen "x", row, 36		'going into the SANC panel to get case info     
	Else 
		EMWriteScreen "x", 7, 36		'going into the SANC panel to get case info
	End if 
	
	transmit
	EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets added to priv case list
	If PRIV_check = "PRIV" then
		priv_case_list = priv_case_list & "|" & MAXIS_case_number
		SET objRange = objExcel.Cells(excel_row, 1).EntireRow
		objRange.Delete				'row gets deleted since it will get added to the priv case list at end of script in col 20
		'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
		Do
			back_to_self
			EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
			If SELF_screen_check <> "SELF" then PF3
		LOOP until SELF_screen_check = "SELF"
		EMWriteScreen "________", 18, 43		'clears the case number
		transmit
	END IF

	'For all of the cases that aren't privileged...
    EMReadScreen ERRR_panel_check, 4, 2, 52         'Ensuring that there are no errors on the case. If they are the client inforamiton will not input.
    If ERRR_panel_check = "ERRR" then transmit
	EMReadScreen memb_number, 2, 4, 12		'reading member number

	'STAT DISA PORTION
	Call navigate_to_MAXIS_screen("STAT", "DISA")
	EMWriteScreen memb_number, 20, 76				'enters member number
	transmit
	'Reading the disa dates
	EMReadScreen disa_start_date, 10, 6, 47			'reading DISA dates
	EMReadScreen disa_end_date, 10, 6, 69
	disa_start_date = Replace(disa_start_date," ","/")		'cleans up DISA dates
	disa_end_date = Replace(disa_end_date," ","/")
	disa_dates = trim(disa_start_date) & " - " & trim(disa_end_date)
	If disa_dates = "__/__/____ - __/__/____" then disa_dates = "NO DISA INFO"
	
    ObjExcel.Cells(excel_row, 5).Value = memb_number		'adds client member number to Excel list
	ObjExcel.Cells(excel_row, 6).Value = disa_dates			'adds disa dates to Excel list
    excel_row = excel_row + 1
	STATS_counter = STATS_counter + 1
LOOP UNTIL objExcel.Cells(excel_row, 1).Value = ""	'Loops until there are no more cases in the Excel list

'Creating the list of privileged cases and adding to the spreadsheet
prived_case_array = split(priv_case_list, "|")
excel_row = 2				'establishes the row to start writing the PRIV cases to

FOR EACH MAXIS_case_number in prived_case_array
	objExcel.cells(excel_row, 7).value = MAXIS_case_number		'inputs cases into Excel
	excel_row = excel_row + 1								'increases the row
NEXT

'------------------------------Post MAXIS coding-----------------------------------------------------------------------------
'Query date/time/runtime info
ObjExcel.Cells(1, 8).Value = "Query date and time:"	'Goes back one, as this is on the next row
objExcel.Cells(1, 8).Font.Bold = TRUE
ObjExcel.Cells(2, 8).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
objExcel.Cells(2, 8).Font.Bold = TRUE
ObjExcel.Cells(1, 9).Value = now
ObjExcel.Cells(2, 9).Value = timer - query_start_time

FOR i = 1 to 9		'formatting the cells'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT
'Logging usage stats
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Please review the list generated.")
