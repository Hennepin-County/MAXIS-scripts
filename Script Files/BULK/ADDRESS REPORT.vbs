'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "BULK - ADDRESS REPORT.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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

BeginDialog x_dlg, 0, 0, 176, 140, "x1 Number"
  EditBox 55, 45, 65, 15, x_number
  CheckBox 20, 65, 140, 10, "Check here to run for the entire county.", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 30, 115, 50, 15
    CancelButton 85, 115, 50, 15
  Text 10, 15, 155, 25, "Please enter the x1 number of the caseload you wish to check (NOTE: please enter the entire 7-digit number):"
  Text 20, 80, 145, 25, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
EndDialog

EMConnect ""

CALL check_for_MAXIS(false)

CALL find_variable("User: ", worker_number, 7)

'Shows dialog
Dialog x_dlg
If buttonpressed = cancel then stopscript

call check_for_MAXIS(false)

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else
	x1s_from_dialog = split(worker_number, ",")	'Splits the worker array based on commas

	'Need to add the worker_county_code to each one
	For each x1_number in x1s_from_dialog
		If worker_array = "" then
			worker_array = worker_county_code & trim(replace(ucase(x1_number), worker_county_code, ""))		'replaces worker_county_code if found in the typed x1 number
		Else
			worker_array = worker_array & ", " & worker_county_code & trim(replace(ucase(x1_number), worker_county_code, "")) 'replaces worker_county_code if found in the typed x1 number
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

objExcel.Cells(1, 1).Value = "WORKER NUMBER"
objExcel.Cells(1, 1).Font.Bold = True
objExcel.Cells(1, 2).Value = "CASE NUMBER"
objExcel.Cells(1, 2).Font.Bold = True
objExcel.Cells(1, 3).Value = "APPLICANT NAME"
objExcel.Cells(1, 3).Font.Bold = True
objExcel.Cells(1, 4).Value = "ADDRESS LINE 1"
objExcel.Cells(1, 4).Font.Bold = True
objExcel.Cells(1, 5).Value = "ADDRESS LINE 2"
objExcel.Cells(1, 5).Font.Bold = True
objExcel.Cells(1, 6).Value = "CITY"
objExcel.Cells(1, 6).Font.Bold = True
objExcel.Cells(1, 7).Value = "STATE"
objExcel.Cells(1, 7).Font.Bold = True
objExcel.Cells(1, 8).Value = "ZIP CODE"
objExcel.Cells(1, 8).Font.Bold = True

excel_row = 2

For each worker in worker_array
	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
	Call navigate_to_screen("rept", "actv")
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

				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
				'If trim(case_number) <> "" and instr(all_case_numbers_array, case_number) <> 0 then exit do
				'all_case_numbers_array = trim(all_case_numbers_array & " " & case_number)

				If case_number = "        " then exit do			'Exits do if we reach the end

				ObjExcel.Cells(excel_row, 1).Value = worker
				ObjExcel.Cells(excel_row, 2).Value = case_number
				ObjExcel.Cells(excel_row, 3).Value = client_name
				excel_row = excel_row + 1

				MAXIS_row = MAXIS_row + 1
				case_number = ""			'Blanking out variable
			Loop until MAXIS_row = 19
			PF8
			
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
next

excel_row = 2
Do
	'Assign case number from Excel
	case_number = ObjExcel.Cells(excel_row, 2)
	
	'Exiting if the case number is blank
	If case_number = "" then exit do

	'Navigate to stat/addr to grab address
	call navigate_to_MAXIS_screen("STAT", "ADDR")
	EMReadScreen priv_check, 4, 2, 50
	If priv_check = "SELF" then 
		objExcel.Cells(excel_row, 4) = "Privileged"
	Else
		EMReadScreen addr_line_1, 22, 6, 43
		EMReadScreen addr_line_2, 22, 7, 43
		EMReadScreen city, 15, 8, 43
		EMReadScreen State, 2, 8, 66
		EMReadScreen Zip_code, 5, 9, 43
		addr_line_1 = replace(addr_line_1, "_", "")
		addr_line_2 = replace(addr_line_2, "_", "")
		city = replace(city, "_", "")
		State = replace(State, "_", "")
		Zip_code = replace(Zip_code, "_", "")
		objExcel.Cells(excel_row, 4) = addr_line_1
		objExcel.Cells(excel_row, 5) = addr_line_2
		objExcel.Cells(excel_row, 6) = city
		objExcel.Cells(excel_row, 7) = State
		objExcel.Cells(excel_row, 8) = Zip_code
	End IF
	
	'Clearing variables for next loop.
	addr_line_1 = ""
	addr_line_2 = ""
	city = ""
	State = ""
	Zip_code = ""
	
	excel_row = excel_row + 1

Loop until case_number = ""
	
	
FOR i = 1 to 8
	objExcel.Columns(i).AutoFit()
NEXT
	
objExcel.Visible = True

script_end_procedure("Success!!")
