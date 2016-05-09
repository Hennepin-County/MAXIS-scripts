'-------------------------->>>>>>>>>>>>>>>>>>>>INFO ABOUT THIS TEMPLATE<<<<<<<<<<<<<<<<<<<<<<<------------------------------------
'This template is a generic form for building custom bulk scripts for MAXIS that collect a list of cases from a rept, then check various panels
'to determine further case information, and if the case meets the desired criteria, add it to an excel list.
'Because the master case list is pulled into an array, rather than to the excel file, time savings should be realized when
'using this format to search a caseload for specific case types.
'The user will need to add a significant amount of coding to this template to have a functional script.

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - TEMPLATE.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 13                      'manual run time in seconds
STATS_denomination = "C"       							'C is for each CASE
'END OF stats block==============================================================================================

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

'Defining classes-----------------------------
'This class holds case-specific data, replace example with any needed info for your bulk list
Class case_attributes
	public case_number
	public worker_number
	public example
END Class

'THE SCRIPT-------------------------------------------------------------------------

'Determining specific county for multicounty agencies...
call get_county_code

'Connects to BlueZone
EMConnect ""

'Creating the dialog

BeginDialog Dialog1, 0, 0, 218, 120, "Pull REPT data into Excel dialog"
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

'Shows dialog
Dialog
If buttonpressed = cancel then stopscript

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Checking for MAXIS
Call check_for_password(false)

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, worker_county_code)
Else
	x1s_from_dialog = split(worker_number, ",")	'Splits the worker array based on commas

	'formatting array
	For each x1_number in x1s_from_dialog
		If worker_array = "" then
			worker_array = trim(x1_number)		'replaces worker_county_code if found in the typed x1 number
		Else
			worker_array = worker_array & ", " & trim(ucase(x1_number)) 'replaces worker_county_code if found in the typed x1 number
		End if
	Next
	'Split worker_array
	worker_array = split(worker_array, ", ")
End if

'Define the case arrays and set any counters to begin case collection

dim case_array() 'This is the array that will hold all cases collected'
case_count = 0 'This will be used to count each case added to the array'



'This template is based around REPT/ACTV, but can be adjusted for other reports by modifying this section of code
'First, we check REPT/ACTV.  Must be done on ACTIVE and CAPER checks'
For each worker in worker_array
	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
	Call navigate_to_MAXIS_screen("rept", "actv")
	EMWriteScreen worker, 21, 13
	transmit
	EMReadScreen user_worker, 7, 21, 71		'
	EMReadScreen p_worker, 7, 21, 13
	IF user_worker = p_worker THEN PF7		'If the user is checking their own REPT/ACTV, the script will back up to page 1 of the REPT/ACTV

	'Skips workers with no info
	EMReadScreen has_content_check, 1, 7, 8
	If has_content_check <> " " then

		'Grabbing each case number on screen
		Do
			'Set variable for next do...loop
			MAXIS_row = 7

			'Checking for the last page of cases.
			EMReadScreen last_page_check, 21, 24, 2	'because on REPT/ACTV it displays right away, instead of when the second F8 is sent
			Do
				EMReadScreen case_number, 8, MAXIS_row, 12		'Reading case number
				EMReadScreen client_name, 21, MAXIS_row, 21		'Reading client name
				EMReadScreen example, 8, MAXIS_row, 42		'Reading example criteria, update variable and coordinates with the data of your choice

				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
				If trim(case_number) <> "" and instr(all_case_numbers_array, case_number) <> 0 then exit do
				all_case_numbers_array = trim(all_case_numbers_array & " " & case_number)

				If case_number = "        " then exit do			'Exits do if we reach the end

				'Using if...thens to decide if a case should be added to the array.  Replace case_number example with your desired case criteria,
				'such as SNAP_status = "A", etc.
				If case_number <> "        " then
					redim preserve case_array(case_count) 'Resize the array for total cases'
					set case_array(case_count) = new case_attributes 'This sets the value of the array at this location to be the case_attributes object, which holds your custom properties'
					case_array(case_count).case_number = case_number
					case_array(case_count).worker_number = worker
					case_array(case_count).example = example 'REPLACE THIS WITH AS MANY case_attibutes properties as you need to read from the REPT
					case_count = case_count+1 'add 1 to the case count'
				END IF

				MAXIS_row = MAXIS_row + 1
				case_number = ""			'Blanking out variable
				STATS_counter = STATS_counter + 1   'adds one instance to the stats counter
			Loop until MAXIS_row = 19
			PF8
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
Next 'This closes out the worker_array loop, and moves to the next worker in worker_array

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True
Set objWorkbook = objExcel.ActiveWorkbook

'Add a worksheet for your cases, label the columns'
ObjExcel.Worksheets.Add().Name = "Example Worksheet"
ObjExcel.Cells(1, 1).Value = "WORKER"
objExcel.Cells(1, 1).Font.Bold = TRUE
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
objExcel.Cells(1, 2).Font.Bold = TRUE
ObjExcel.Cells(1, 3).Value = "EXAMPLE"
ObjExcel.Cells(1, 3).Font.Bold = TRUE

excel_row = 2 '2, because row 1 has column headings

'THE FOLLOWING SECTION WILL LOOP THROUGH EVERY CASE THAT WAS SELECTED FROM REPT/ACTV
'Enter logic to read the needed maxis screens for finding further case criteria'
'And add the cases you want to your spreadsheet
For current_case = 0 to ubound(case_array)
	case_number = case_array(current_case).case_number 'passing the case_number back to the variable used by functions'

	call navigate_to_MAXIS_screen("CASE", "CURR") 'This could be any desired MAXIS screen, only an example
	this_is_a_case_to_add = true 'remove this, it is only for testing.
	'You will enter all your logic for looking up maxis info HERE

	IF this_is_a_case_to_add = true THEN 'Change this variable to whatever it is you are looking for
					ObjExcel.cells(excel_row, 1).value = case_array(current_case).worker_number 'Column 1 is worker number
					ObjExcel.Cells(excel_row, 2).value = case_array(current_case).case_number 'Column 2 is case number'
					ObjExcel.Cells(excel_row, 3).value = case_array(current_case).example 'Enter any criteria you defined here and in subsequent rows'
					excel_row = excel_row + 1 'move to the next row

	END IF
Next 'This closes the current_case FOR loop and moves to the next case


'Query stats, adjust the location if you have more than 5 columns of data.

objExcel.Cells(1, 6).Font.Bold = TRUE
objExcel.Cells(2, 6).Font.Bold = TRUE
ObjExcel.Cells(1, 6).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, 7).Value = now
ObjExcel.Cells(2, 6).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, 7).Value = timer - query_start_time

'Autofitting columns
For col_to_autofit = 1 to 7
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

'Logging usage stats
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("")
