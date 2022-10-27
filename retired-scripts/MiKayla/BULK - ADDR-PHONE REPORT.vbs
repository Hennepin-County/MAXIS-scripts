'Required for statistical purposes===============================================================================
name_of_script = "BULK - ADDR-PHONE REPORT.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 36                               'manual run time in seconds
STATS_denomination = "I"       'I is for each Item
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

call changelog_update("08/25/2017", "Initial version.", "MiKayla Handley, Hennepin")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'---------------------------------------------------------------------------------------------SCRIPT
EMConnect ""
CALL check_for_MAXIS(True)
Call find_MAXIS_worker_number(x_number)

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 251, 120, "Pull ADDR data into Excel"
  	EditBox 75, 40, 160, 15, worker_number
  	CheckBox 10, 70, 150, 10, "Check here to run this query county-wide.", all_workers_check
  		ButtonGroup ButtonPressed
    	OkButton 140, 100, 50, 15
    CancelButton 195, 100, 50, 15
  	Text 10, 45, 65, 10, "Worker(s) to check:"
  	Text 10, 85, 235, 10, "NOTE: running queries county-wide may take several hours to complete"
  	Text 10, 10, 170, 10, "Enter workers' x1 numbers, separated by a comma."
  	Text 10, 25, 100, 10, "EX: X_ _ _ _ _ _, X_ _ _ _ _ _"
  	GroupBox 5, 0, 235, 60, ""
EndDialog
'Shows dialog
Do
	Do
		err_msg = ""
		Dialog Dialog1
		If buttonpressed = cancel then stopscript
		If (all_workers_check = 0 AND worker_number = "") then err_msg = err_msg & vbNewLine & "* Please enter at least one worker number." 'allows user to select the all workers check, and not have worker number be ""
		If (all_workers_check = 1 AND trim(worker_number) <> "") then err_msg = err_msg & vbNewLine & "* Please enter x numbers OR the county-wide query."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
  	LOOP until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until check_for_password(are_we_passworded_out) = False		'loops until user is password-ed out

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else
	x1s_from_dialog = split(x_number, ",")	'Splits the worker array based on commas

	'Need to add the worker_county_code to each one
	For each x1_number in x1s_from_dialog
		If worker_array = "" then
			worker_array = x_number
		Else
			worker_array = worker_array & ", " & x_number 'replaces worker_county_code if found in the typed x1 number
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

'Creating columns
objExcel.Cells(1, 1).Value = "WORKER NUMBER"
objExcel.Cells(1, 2).Value = "CASE NUMBER"
objExcel.Cells(1, 3).Value = "APPLICANT NAME"
objExcel.Cells(1, 4).Value = "ADDRESS LINE 1"
objExcel.Cells(1, 5).Value = "ADDRESS LINE 2"
objExcel.Cells(1, 6).Value = "CITY"
objExcel.Cells(1, 7).Value = "STATE"
objExcel.Cells(1, 8).Value = "ZIP CODE"
objExcel.Cells(1, 9).Value = "MAILING ADDRESS LINE 1"
objExcel.Cells(1, 10).Value = "MAILING ADDRESS LINE 2"
objExcel.Cells(1, 11).Value = "MAILING CITY"
objExcel.Cells(1, 12).Value = "MAILING STATE"
objExcel.Cells(1, 13).Value = "MAILING ZIP CODE"
objExcel.Cells(1, 14).Value = "HOMELESS"
objExcel.Cells(1, 15).Value = "PHONE"
objExcel.Cells(1, 16).Value = "PHONE II"
objExcel.Cells(1, 17).Value = "PHONE III"


FOR i = 1 to 26		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 	'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'

NEXT

excel_row = 2

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
				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 12		'Reading case number
				EMReadScreen client_name, 21, MAXIS_row, 21		'Reading client name

				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
				'If trim(MAXIS_case_number) <> "" and instr(all_case_numbers_array, MAXIS_case_number) <> 0 then exit do
				'all_case_numbers_array = trim(all_case_numbers_array & " " & MAXIS_case_number)

				If MAXIS_case_number = "        " then exit do			'Exits do if we reach the end

				ObjExcel.Cells(excel_row, 1).Value = worker
				ObjExcel.Cells(excel_row, 2).Value = MAXIS_case_number
				ObjExcel.Cells(excel_row, 3).Value = client_name
				excel_row = excel_row + 1

				MAXIS_row = MAXIS_row + 1
				MAXIS_case_number = ""			'Blanking out variable
			Loop until MAXIS_row = 19
			PF8

		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
next

excel_row = 2
Do
	'Assign case number from Excel
	MAXIS_case_number = ObjExcel.Cells(excel_row, 2)

	'Exiting if the case number is blank
	If MAXIS_case_number = "" then exit do

	'Navigate to stat/addr to grab address
	call navigate_to_MAXIS_screen("STAT", "ADDR")
	EMReadScreen priv_check, 4, 2, 50
	If priv_check = "SELF" then
		objExcel.Cells(excel_row, 4) = "Privileged"
	Else
		'Reading and cleaning up Residence address
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
		'Reading homeless code
		EMReadScreen homeless_code, 1, 10, 43
		'Reading and cleaning up mailing address
		EMReadScreen mailing_addr_line_1, 22, 13, 43
		EMReadScreen mailing_addr_line_2, 22, 14, 43
		EMReadScreen mailing_city, 15, 15, 43
		EMReadScreen mailing_State, 2, 16, 43
		EMReadScreen mailing_Zip_code, 5, 16, 52
		mailing_addr_line_1 = replace(mailing_addr_line_1, "_", "")
		mailing_addr_line_2 = replace(mailing_addr_line_2, "_", "")
		mailing_city = replace(mailing_city, "_", "")
		mailing_State = replace(mailing_State, "_", "")
		mailing_Zip_code = replace(mailing_Zip_code, "_", "")

		EMReadScreen phone_home, 14, 17, 45
		EMReadScreen phone_homeII, 14, 18, 45
		EMReadScreen phone_homeIII, 14, 19, 45
		phone_home = replace(phone_home, "_", "")
		phone_homeII = replace(phone_homeII, "_", "")
		phone_homeIII = replace(phone_homeIII, "_", "")
		phone_home = replace(phone_home, ")", "")
		phone_homeII = replace(phone_homeII, ")", "")
		phone_homeIII = replace(phone_homeIII, ")", "")
		phone_home = replace(phone_home, " ", "")
		phone_homeII = replace(phone_homeII, " ", "")
		phone_homeIII = replace(phone_homeIII, " ", "")
		'phone_home = replace(mailing_addr_line_1, "_", "") want to format to xxx-xxx-xxxx
		'phone_homeII = replace(mailing_addr_line_1, "_", "")
		'phone_homeIII = replace(mailing_addr_line_1, "_", "")
		'Writing both addresses into excel
		objExcel.Cells(excel_row, 4) = addr_line_1
		objExcel.Cells(excel_row, 5) = addr_line_2
		objExcel.Cells(excel_row, 6) = city
		objExcel.Cells(excel_row, 7) = State
		objExcel.Cells(excel_row, 8) = Zip_code
		objExcel.Cells(excel_row, 9) = mailing_addr_line_1
		objExcel.Cells(excel_row, 10) = mailing_addr_line_2
		objExcel.Cells(excel_row, 11) = mailing_city
		objExcel.Cells(excel_row, 12) = mailing_State
		objExcel.Cells(excel_row, 13) = mailing_Zip_code
		objExcel.Cells(excel_row, 14) = homeless_code
		objExcel.Cells(excel_row, 15) = phone_home
		objExcel.Cells(excel_row, 16) = phone_homeII
		objExcel.Cells(excel_row, 17) = phone_homeIII
	End IF

	'Clearing variables for next loop.
	addr_line_1 = ""
	addr_line_2 = ""
	city = ""
	State = ""
	Zip_code = ""
	mailing_addr_line_1 = ""
	mailing_addr_line_2 = ""
	mailing_city = ""
	mailing_State = ""
	mailing_Zip_code = ""
	homeless_code = ""
	phone_home = ""
	phone_homeII = ""
	phone_homeIII = ""


	excel_row = excel_row + 1
	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
Loop until MAXIS_case_number = ""

'formatting excel columns to fit
FOR i = 1 to 14
	objExcel.Columns(i).AutoFit()
NEXT

'making excel document visible.
objExcel.Visible = True

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success!!")
