'Required for statistical purposes===============================================================================
name_of_script = "ADMIN - ADDRESS REPORT.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 40                               'manual run time in seconds
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
'END FUNCTIONS LIBRARY BLOCK===================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("02/22/2024", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------------------------
EMConnect ""
get_county_code 'Checks for county info from global variables, or asks if it is not already defined.
CALL check_for_MAXIS(True)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 105, "Address Report Dialog"
  GroupBox 10, 5, 250, 55, "Using the Address Report Script:"
  Text 20, 20, 235, 35, "This script will pull the resident and mailing addresses for residents open on SNAP and/or Cash programs effort to ensure that residents are getting their mail after an initial need to get an EBT card through a Hennepin County building address."
  Text 15, 70, 60, 10, "Worker number(s):"
  EditBox 80, 65, 180, 15, worker_number
  CheckBox 15, 90, 135, 10, "Check here to process for all workers.", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 155, 85, 50, 15
    CancelButton 210, 85, 50, 15
EndDialog

Do
	Do
  		err_msg = ""
  		dialog Dialog1
  		cancel_without_confirmation
  		If trim(worker_number) = "" and all_workers_check = 0 then err_msg = err_msg & vbNewLine & "* Select a worker number(s) or all cases."
  		If trim(worker_number) <> "" and all_workers_check = 1 then err_msg = err_msg & vbNewLine & "* Select a worker number(s) or all cases, not both options."
  	  	If err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
  	LOOP until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call check_for_MAXIS(False)

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else
	x1s_from_dialog = split(worker_number, ", ")	'Splits the worker array based on commas
	'Need to add the worker_county_code to each one
	For each x1_number in x1s_from_dialog
		If worker_array = "" then
			worker_array = trim(ucase(x1_number))		'replaces worker_county_code if found in the typed x1 number
		Else
			worker_array = worker_array & "," & trim(ucase(x1_number)) 'replaces worker_county_code if found in the typed x1 number
		End if
	Next
	'Split worker_array
	worker_array = split(worker_array, ",")
End if

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Creating columns
objExcel.Cells(1, 1).Value = "X NUMBER"
objExcel.Cells(1, 2).Value = "CASE #"
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

FOR i = 1 to 14		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

excel_row = 2
back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
Call navigate_to_MAXIS_screen("REPT", "ACTV")

For each worker in worker_array
    Call write_value_and_transmit(worker, 21, 13)
	EMReadScreen has_content_check, 1, 7, 8     	'Skips workers with no info
	If has_content_check <> " " then
		Do                          
			MAXIS_row = 7  			'Set variable for next do...loop
			EMReadScreen last_page_check, 21, 24, 2		'Checking for the last page of cases because on REPT/ACTV it displays right away, instead of when the second F8 is sent
			Do
				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 12		'Reading case number
				MAXIS_case_number = trim(MAXIS_case_number)
                If MAXIS_case_number = "" then exit do			'Exits do if we reach the end
                capture_case = False 
                EMReadScreen cash_program, 1, MAXIS_row, 54
                EmReadScreen cash_program_2, 1, MAXIS_row, 59
                EMReadScreen SNAP_program, 1, MAXIS_row, 61
                
                If cash_program = "A" then capture_case = True 
                If cash_program_2 = "A" then capture_case = True 
                If SNAP_program = "A" then capture_case = True 

                If capture_case = True then 
                    EMReadScreen client_name, 21, MAXIS_row, 21		'Reading client name    
				    ObjExcel.Cells(excel_row, 1).Value = worker
				    ObjExcel.Cells(excel_row, 2).Value = MAXIS_case_number
				    ObjExcel.Cells(excel_row, 3).Value = trim(client_name)
				    excel_row = excel_row + 1
                End if 
				MAXIS_row = MAXIS_row + 1
				MAXIS_case_number = ""			'Blanking out variable
			Loop until MAXIS_row = 19
			PF8
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
Next

'Filling in STAT/ADDR information for SNAP and/or Cash programs that are in an active status. 
excel_row = 2
Do
	MAXIS_case_number = ObjExcel.Cells(excel_row, 2)    'Assign case number from Excel
	If MAXIS_case_number = "" then exit do  	'Exiting if the case number is blank
	Call navigate_to_MAXIS_screen_review_PRIV("STAT", "ADDR", is_this_priv)  'Navigate to stat/addr to grab address - checking for PRIV
    If is_this_priv = True then
        objExcel.Cells(excel_row, 4).Value = "Privileged"
	Else
        Call access_ADDR_panel("READ", notes_on_address, resi_line_one, resi_line_two, resi_street_full, resi_city, resi_state, resi_state, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city, mail_state, mail_zip, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)
		'Writing both addresses into excel
		objExcel.Cells(excel_row, 4) = resi_line_one
		objExcel.Cells(excel_row, 5) = resi_line_two
		objExcel.Cells(excel_row, 6) = resi_city
		objExcel.Cells(excel_row, 7) = resi_state
		objExcel.Cells(excel_row, 8) = resi_state
		objExcel.Cells(excel_row, 9) = mail_line_one
		objExcel.Cells(excel_row, 10) = mail_line_two
		objExcel.Cells(excel_row, 11) = mail_city
		objExcel.Cells(excel_row, 12) = mail_state
		objExcel.Cells(excel_row, 13) = mail_zip
		objExcel.Cells(excel_row, 14) = addr_homeless
	End if 
	excel_row = excel_row + 1
	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
Loop until MAXIS_case_number = ""

'formatting excel columns to fit
FOR i = 1 to 14
	objExcel.Columns(i).AutoFit()
NEXT

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure_with_error_reporting("Success! Your list has been generated.")
