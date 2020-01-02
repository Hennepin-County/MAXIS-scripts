'Required for statistical purposes===============================================================================
name_of_script = "BULK - SWKR LIST GENERATOR.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 45                      'manual run time in seconds
STATS_denomination = "C"       						 'C is for each CASE
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

'The script----------------------------------------------------------------------------------------------------------------------------------
EMConnect ""

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 136, 100, "Get SWKE Information"
  EditBox 65, 10, 60, 15, x_number
  CheckBox 20, 30, 95, 10, "Check here for all workers", all_workers_check
  CheckBox 20, 60, 100, 10, "Restart from previous list.", restart_checkbox
  ButtonGroup ButtonPressed
    OkButton 20, 80, 50, 15
    CancelButton 75, 80, 50, 15
  Text 55, 45, 30, 10, "***OR***"
  Text 5, 15, 60, 10, "Worker to check:"
EndDialog
Do 
    err_msg = ""
    dialog Dialog1 
    cancel_without_confirmation
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Starting the query start time (for the query runtime at the end)
 query_start_time = timer
'IF x_number = "" THEN CALL find_variable("User: ", x_number, 7)

If restart_checkbox = 0 then
    'Opening the Excel file
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = True
    Set objWorkbook = objExcel.Workbooks.Add() 
    objExcel.DisplayAlerts = True
    
	'FORMATS THE EXCEL SPREADSHEET WITH THE HEADERS, AND SETS THE COLUMN WIDTH
	ObjExcel.Cells(1, 1).Value = "Basket"
	ObjExcel.Cells(1, 2).Value = "Case #"
	ObjExcel.Cells(1, 3).Value = "Name"
	ObjExcel.Cells(1, 4).Value = "SWKR name"
	ObjExcel.Cells(1, 5).Value = "Address 1"
	ObjExcel.Cells(1, 6).Value = "Address 2"
	ObjExcel.Cells(1, 7).Value = "City"
	ObjExcel.Cells(1, 8).Value = "State"
	ObjExcel.Cells(1, 9).Value = "Zip"
	ObjExcel.Cells(1, 10).Value = "Phone number"
	ObjExcel.Cells(1, 11).Value = "Send Notice?"
	
	FOR i = 1 to 11		'formatting the cells'
		objExcel.Cells(1, i).Font.Bold = True		'bold font'
		ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
		objExcel.Columns(i).AutoFit()				'sizing the columns'
	NEXT

    excel_row = 2
    back_to_SELF
    
    'If all workers are selected, the script will open the worker list stored on the shared drive, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
    If all_workers_check = 1 then
    	CALL create_array_of_all_active_x_numbers_in_county(x_array, "27")
    Else
    	IF len(x_number) > 3 THEN 
    		x_array = split(x_number, ", ")
    	ELSE		
    		x_array = split(x_number)
    	END IF
    End if
	
    For each worker in x_array
    	'Getting to ACTV, if ACTV is the selected option
    	Call navigate_to_MAXIS_screen("rept", "actv")
    	EMWriteScreen worker, 21, 13
    	transmit
    
        'setting variables for first run through
        rept_row = 7
        DO
        	EMReadScreen last_page, 21, 24, 2										'checking to see if this is the last page, if it is the loop can end.
        	DO
        		EMReadScreen MAXIS_case_number, 8, rept_row, 12						'reading the case numbers from rept/actv
        		MAXIS_case_number = trim(MAXIS_case_number)
				If MAXIS_case_number = "" then exit do
        		
        		EMReadScreen client_name, 21, rept_row, 21						'grabbing client name
        		
    			'Inputting variables in the spreadsheet
    			objExcel.Cells(excel_row, 1).Value = worker					'adding read variables to the spreadsheet
    			objExcel.Cells(excel_row, 2).Value = MAXIS_case_number			'adding read variables to the spreadsheet
    			objExcel.Cells(excel_row, 3).Value = trim(client_name)				'adding read variables to the spreadsheet
        		
				excel_row = excel_row + 1
        		rept_row = rept_row + 1
        	LOOP UNTIL rept_row = 19								'looping until the script reads through the bottom of the page.
        	PF8														'pf8 navigates to next page of ACTV
        	rept_row = 7											'resetting the row to the top of the page.
        	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
        LOOP UNTIL last_page = "THIS IS THE LAST PAGE"
    Next 	
    
    Excel_row = 2
Else 
    'dialog and dialog DO...Loop	
    Do
    	Do
    		'The dialog is defined in the loop as it can change as buttons are pressed 
    		Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 221, 50, "Select the ABAWD pull cases into Excel file."
    			ButtonGroup ButtonPressed
    			PushButton 175, 10, 40, 15, "Browse...", select_a_file_button
    			OkButton 110, 30, 50, 15
    			CancelButton 165, 30, 50, 15
    			EditBox 5, 10, 165, 15, file_selection_path
    		EndDialog
    		err_msg = ""
    		Dialog Dialog1 
    		cancel_without_confirmation
        	If ButtonPressed = select_a_file_button then
        		If file_selection_path <> "" then 'This is handling for if the BROWSE button is pushed more than once'
        			objExcel.Quit 'Closing the Excel file that was opened on the first push'
        			objExcel = "" 	'Blanks out the previous file path'
        		End If
        		call file_selection_system_dialog(file_selection_path, ".xlsx") 'allows the user to select the file'
        	End If
        	If file_selection_path = "" then err_msg = err_msg & vbNewLine & "Use the Browse Button to select the file that has your client data"
        	If err_msg <> "" Then MsgBox err_msg
        Loop until err_msg = ""
        If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
        If err_msg <> "" Then MsgBox err_msg
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in
	
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 126, 50, "Select the excel row to restart"
      EditBox 75, 5, 40, 15, excel_row_to_restart
      ButtonGroup ButtonPressed
        OkButton 10, 25, 50, 15
        CancelButton 65, 25, 50, 15
      Text 10, 10, 60, 10, "Excel row to start:"
    EndDialog
	do 
		dialog dialog1
		cancel_without_confirmation
	    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	Loop until are_we_passworded_out = false					'loops until user passwords back in
	
    excel_row = excel_row_to_restart
ENd if 

'NOW THE SCRIPT IS CHECKING STAT/AREP FOR EACH CASE.----------------------------------------------------------------------------------------------------
back_to_self

Do 
	MAXIS_case_number = ObjExcel.Cells(excel_row, 2).Value
	If trim(MAXIS_case_number) = "" then exit do

	call navigate_to_MAXIS_screen("STAT", "SWKR")
	EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets added to priv case list
	If PRIV_check = "PRIV" then
		ObjExcel.Cells(excel_row, 4).Value = "PRIV case. Cannot access."
		'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
		Do
			back_to_self
			EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
			If SELF_screen_check <> "SELF" then PF3
		LOOP until SELF_screen_check = "SELF"
		EMWriteScreen "________", 18, 43		'clears the case number
		transmit
	Else 
	    'NAVIGATES TO SWKR, and adds the applicable inforamtion to the spreadsheet
	    EMReadScreen SWKR_name, 34, 6, 32
	    EMReadScreen addr_one, 22, 8, 32
	    EMReadScreen addr_two, 22, 9, 32
	    EMReadScreen city, 15, 10, 32
	    EMReadScreen state, 2, 10, 54
	    EMReadScreen zip_code, 7, 10, 63
	    EMReadScreen phone_number, 26, 12, 32
	    If phone_number = "( ___ ) ___ ____ Ext: ____" then phone_number = ""
	    EMReadScreen send_notice, 1, 15, 63
	    
	    ObjExcel.Cells(excel_row, 4).Value = replace(swkr_name, "_", "")
	    ObjExcel.Cells(excel_row, 5).Value = replace(addr_one, "_", "")
	    ObjExcel.Cells(excel_row, 6).Value = replace(addr_two, "_", "")
	    ObjExcel.Cells(excel_row, 7).Value = replace(city, "_", "")
	    ObjExcel.Cells(excel_row, 8).Value = replace(state, "_", "")
	    ObjExcel.Cells(excel_row, 9).Value = replace(zip_code, "_", "")
	    ObjExcel.Cells(excel_row, 10).Value = replace(phone_number, "_", "")
	    ObjExcel.Cells(excel_row, 11).Value = replace(send_notice, "_", "")
	End if 

	excel_row = excel_row + 1 'setting up the script to check the next row.
loop

FOR i = 1 to 11		'formatting the cells'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Your list has been created.")
