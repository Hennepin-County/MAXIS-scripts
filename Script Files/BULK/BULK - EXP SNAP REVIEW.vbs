worker_county_code = "x127"

'Required for statistical purposes===============================================================================
name_of_script = "BULK - EXP SNAP REVIEW.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 0                      'manual run time in seconds
STATS_denomination = "C"       							'C is for each CASE
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

'DIALOGS-------------------------------------------------------------------------------------------------------------
BeginDialog EXP_SNAP_review_dialog, 0, 0, 286, 185, "EXP SNAP review "
  EditBox 75, 10, 205, 15, worker_number
  CheckBox 25, 50, 140, 10, "Check here to run this query county-wide ", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 175, 45, 50, 15
    CancelButton 230, 45, 50, 15
  Text 5, 15, 65, 10, "Worker(s) to check:"
  Text 10, 90, 270, 20, "This script will create a list of cases that should be reviewed for expedited SNAP eligibilty from REPT/PND1 and REPT/PND2."
  Text 5, 30, 275, 10, "Enter all 7 digits of your workers' x1 numbers (ex: x######), separated by a comma."
  GroupBox 5, 75, 280, 105, "BULK - Expedited SNAP review"
  Text 10, 115, 270, 30, "* The REPT/PND1 list of cases will include ALL cases that do not have a case note that identifies the case as not expedited. This includes cases that are not pending for SNAP since REPT/PND1 does not make the distinction."
  Text 10, 145, 270, 25, "*The REPT/PND2 list of cases will identify ALL cases that are pending for SNAP (or MFIP if SNAP isn't active) that do not have a case note that identifies the case as not expedited."
EndDialog

'THE SCRIPT-----------------------------------------------------------------------------------------------------------
'Determining specific county for multicounty agencies...
get_county_code

'Connects to BlueZone
EMConnect ""

'Shows dialog
DO 
    Dialog EXP_SNAP_review_dialog
    If buttonpressed = cancel then script_end_procedure("")
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
Loop until are_we_passworded_out = false					'loops until user passwords back in					

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

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

'Sets up the array to store all the information for each client'
Dim PND1_array ()
ReDim PND1_array (5, 0)

'Sets constants for the array to make the script easier to read (and easier to code)'
Const work_num          = 1     	'Each of the case numbers will be stored at this position'		
Const case_num          = 2			
Const clt_name          = 3
Const app_date          = 4
Const days_pending      = 5

For each worker in worker_array
	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
	Call navigate_to_MAXIS_screen("REPT", "PND1")
	EMWriteScreen worker, 21, 13
	transmit

	'Skips workers with no info
	EMReadScreen has_content_check, 8, 7, 3
	If has_content_check <> "        " then
		'Grabbing each case number on screen
		Do
			'Set variable for next do...loop
			MAXIS_row = 7
			Do
				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 3			'Reading case number
				MAXIS_case_number = trim(MAXIS_case_number)
				EMReadScreen client_name, 25, MAXIS_row, 13		'Reading client name
				EMReadScreen appl_date, 8, MAXIS_row, 41		      'Reading application date
				appl_date = replace(appl_date, " ", "/")
				EMReadScreen nbr_days_pending, 4, MAXIS_row, 54		'Reading nbr days pending
				
				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
				If trim(MAXIS_case_number) <> "" and instr(all_case_numbers_array, MAXIS_case_number) <> 0 then exit do
				all_case_numbers_array = trim(all_case_numbers_array & " " & MAXIS_case_number)
				If MAXIS_case_number = "        " then exit do			'Exits do if we reach the end
				
				'Now the script adds all the clients on the excel list into an array
				entry_record = 0
				'Adding client information to the array'
				ReDim Preserve PND1_array(5, entry_record)	'This resizes the array based on the number of rows in the Excel File'
				'The client information is added to the array'
				PND1_array (work_num,   entry_record) = worker
				PND1_array (case_num,	entry_record) = MAXIS_case_number		
				PND1_array (clt_name,  		entry_record) = client_name
				PND1_array (app_date, 			entry_record) = APPL_date
				PND1_array (days_pending,   entry_record) = nbr_days_pending
				
				entry_record = entry_record + 1			'This increments to the next entry in the array'
				MAXIS_row = MAXIS_row + 1		
			Loop until MAXIS_row = 19
			PF8
			EMReadScreen last_page_check, 21, 24, 2
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
next

'Now the script goes into CASENOTE and searches for evidence that EXP screening has
For item = 0 to UBound(PND1_array, 2)
	MAXIS_case_number = Banked_Month_Client_Array(case_num, item)	'Case number is set for each loop
	Do	
		back_to_self
		EMWriteScreen "________", 18, 43
		EMWriteScreen MAXIS_case_number, 18, 43
		Call navigate_to_MAXIS_screen("CASE", "NOTE")
	
		MAXIS_row = 5
		Do 
			EMReadScreen case_note_date, 8, MAXIS_row, 6
			If case_note_date = "" then exit do
			msgbox case_note_date & vbnewline & "appl date: " &  APPL_date
			msgbox "case note date is => then appl date: " & case_note_date => APPL_date
			If case_note_date => APPL_date then 
				EMReadScreen case_note_header, 55, MAXIS_row, 25
				case_note_header = trim(case_note_header)	
				msgbox case_note_header
				IF instr(case_note_header, "client appears expedited") then
					'Deleting the blank results to clean up the spreadsheet
					SET objRange = objExcel.Cells(excel_row, 1).EntireRow
					objRange.Delete
					excel_row = excel_row - 1		'does not advance one row if the case gets deleted
					exit do
				END IF
				MAXIS_row = MAXIS_row + 1
			END IF
		LOOP until case_note_date < APPL_date
		MAXIS_case_number = ""
		excel_row = excel_row + 1
		STATS_counter = STATS_counter + 1
	Loop until MAXIS_case_number = ""
		
msgbox "all done with PND1"		

''PND2 information
''Sets up the array to store all the information for each client'
'Dim PND2_array ()
'ReDim PND2_array (5, 0)
'
''Sets constants for the array to make the script easier to read (and easier to code)'
'Const work_num          = 1     	'Each of the case numbers will be stored at this position'		
'Const case_num          = 2			
'Const clt_name          = 3
'Const app_date          = 4
'Const days_pending      = 5
' 
'For each worker in worker_array
'	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
'	Call navigate_to_MAXIS_screen("REPT", "PND2")
'	EMWriteScreen worker, 21, 13
'	transmit
'
'	'Skips workers with no info
'	EMReadScreen has_content_check, 8, 7, 3
'	If has_content_check <> "        " then
'		'Grabbing each case number on screen
'		Do
'			'Set variable for next do...loop
'			MAXIS_row = 7
'			Do
'				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 5			'Reading case number
'				MAXIS_case_number = trim(MAXIS_case_number)				
'				EMReadScreen client_name, 25, MAXIS_row, 16				'Reading client name
'				EMReadScreen appl_date, 8, MAXIS_row, 38		      	'Reading application date
'				appl_date = replace(appl_date, " ", "/")
'				EMReadScreen nbr_days_pending, 4, MAXIS_row, 49			'Reading nbr days pending
'				EMReadScreen cash_pending, 2, MAXIS_row, 56				'checking for pending MFIP
'				EMReadScreen SNAP_pending, 1, MAXIS_row, 62				'checking for pending SNAP
'				
'				'if more than one application date exists, the case number is not present. This fixes that
'				If trim(MAXIS_case_number) = "" AND trim(client_name) <> "" then 			'if there's a name and no case number
'					EMReadScreen alt_case_number, 8, MAXIS_row - 1, 6				'then it reads the row above
'					MAXIS_case_number = alt_case_number									'restablishes that in this instance, alt case number = case number'
'				END IF
'				
'				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
'				If trim(MAXIS_case_number) <> "" and instr(all_case_numbers_array, MAXIS_case_number) <> 0 then exit do
'				all_case_numbers_array = trim(all_case_numbers_array & " " & MAXIS_case_number)
'				If MAXIS_case_number = "        " then exit do			'Exits do if we reach the end
'				
'				'Now the script adds all the clients on the excel list into an array
'				entry_record = 0
'				'Adding client information to the array'
'				ReDim Preserve PND2_array(5, entry_record)	'This resizes the array based on the number of rows in the Excel File'
'				'The client information is added to the array'
'				PND2_array (work_num,   entry_record) = worker
'				PND2_array (case_num,	entry_record) = MAXIS_case_number		
'				PND2_array (clt_name,  		entry_record) = client_name
'				PND2_array (app_date, 			entry_record) = APPL_date
'				PND2_array (days_pending,   entry_record) = nbr_days_pending
'				
'				entry_record = entry_record + 1			'This increments to the next entry in the array'
'				MAXIS_row = MAXIS_row + 1		
'			Loop until MAXIS_row = 19
'			PF8
'			EMReadScreen last_page_check, 21, 24, 2
'		Loop until last_page_check = "THIS IS THE LAST PAGE"
'	End if
'	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
'next

'Now the script goes into CASENOTE for PND2 cases that have SNAP or MFIP pending, and searches for evidence that EXP screening has
'For item = 0 to UBound(PND2_array, 2)
'	MAXIS_case_number = Banked_Month_Client_Array(case_num, item)	'Case number is set for each loop
'
'	Do
'		back_to_self
'		EMWriteScreen "________", 18, 43
'		EMWriteScreen MAXIS_case_number, 18, 43
'		Call navigate_to_MAXIS_screen("CASE", "NOTE")
'		
'		MAXIS_row = 5
'		Do 
'			EMReadScreen case_note_date, 8, MAXIS_row, 6
'			If case_note_date = "" then exit do
'			msgbox case_note_date & vbnewline & "appl date: " &  APPL_date
'			msgbox "case note date is => then appl date: " & case_note_date => APPL_date
'			If case_note_date => APPL_date then 
'				EMReadScreen case_note_header, 55, MAXIS_row, 25
'				case_note_header = trim(case_note_header)	
'				msgbox case_note_header
'				IF instr(case_note_header, "client appears expedited") then
'					'Deleting the blank results to clean up the spreadsheet
'					SET objRange = objExcel.Cells(excel_row, 1).EntireRow
'					objRange.Delete
'					excel_row = excel_row - 1		'does not advance one row if the case gets deleted
'					exit do
'				END IF
'				MAXIS_row = MAXIS_row + 1
'			END IF
'		LOOP until case_note_date < APPL_date
'		MAXIS_case_number = ""
'		excel_row = excel_row + 1
'		STATS_counter = STATS_counter + 1
'	Loop until MAXIS_case_number = ""
'
'msgbox "all done with PND2"	

Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

Changes name of Excel sheet to "Case information"
ObjExcel.ActiveSheet.Name = "PND1 cases"

'Setting the variable for what's to come
excel_row = 2
'
adding information to the Excel list
ObjExcel.Cells(excel_row, 1).Value = worker
ObjExcel.Cells(excel_row, 2).Value = MAXIS_case_number
ObjExcel.Cells(excel_row, 3).Value = client_name
ObjExcel.Cells(excel_row, 4).Value = appl_date
ObjExcel.Cells(excel_row, 5).Value = abs(nbr_days_pending)
excel_row = excel_row + 1

'add_to_excel = ""			'Blanking out variables		
'MAXIS_case_number = ""	

''Adding another sheet 
''ObjExcel.Worksheets.Add().Name = "PND2 cases"
'
''Setting the first 4 col as worker, case number, name, and APPL date
'ObjExcel.Cells(1, 1).Value = "WORKER"
'ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
'ObjExcel.Cells(1, 3).Value = "NAME"
'ObjExcel.Cells(1, 4).Value = "APPL DATE"
'objExcel.Columns(4).NumberFormat = "mm/dd/yy"	
'ObjExcel.Cells(1, 5).Value = "NBR DAYS PENDING"
'
'FOR i = 1 to 5		'formatting the cells
'	objExcel.Cells(1, i).Font.Bold = True		'bold font'
'	objExcel.Columns(i).AutoFit()				'sizing the columns'
'NEXT	

'''col_to_use is normally used for setting a variable amount of columns. PND1, however, always uses the same amount. I'm setting it as a firm variable here, but this could just as easily include a "col_to_use = col_to_use + 2", like in the PND2 script. -VKC, 01/12/2015
'col_to_use = 10
'
''Query date/time/runtime info
'objExcel.Cells(1, col_to_use - 1).Font.Bold = TRUE
'objExcel.Cells(2, col_to_use - 1).Font.Bold = TRUE
'ObjExcel.Cells(1, col_to_use - 1).Value = "Query date and time:"	'Goes back one, as this is on the next row
'ObjExcel.Cells(1, col_to_use).Value = now
'ObjExcel.Cells(2, col_to_use - 1).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
'ObjExcel.Cells(2, col_to_use).Value = timer - query_start_time
'
''Autofitting columns
'For col_to_autofit = 1 to col_to_use
'	ObjExcel.columns(col_to_autofit).AutoFit()
'Next

'logging usage stats
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Please review the PND1 and PND2 lists for potential EXP SNAP processing.")
