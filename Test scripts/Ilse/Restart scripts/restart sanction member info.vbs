'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - SANCTION MEMBER INFO.vbs"
start_time = timer
STATS_counter = 1                       'sets the stats counter at one
STATS_manualtime = "60"                'manual run time in seconds
STATS_denomination = "C"       			'C is for each CASE
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("05/25/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOGS----------------------------------------------------------------------
BeginDialog ABAWD_report_dialog, 0, 0, 136, 100, "ABAWD report from REPT/ACTV"
  EditBox 65, 10, 60, 15, x_number
  CheckBox 20, 30, 95, 10, "Check here for all workers", all_workers_check
  CheckBox 20, 60, 100, 10, "Restart from previous list.", restart_checkbox
  ButtonGroup ButtonPressed
    OkButton 20, 80, 50, 15
    CancelButton 75, 80, 50, 15
  Text 55, 45, 30, 10, "***OR***"
  Text 5, 15, 60, 10, "Worker to check:"
EndDialog

BeginDialog excel_row_dialog, 0, 0, 126, 50, "Select the excel row to restart"
  EditBox 75, 5, 40, 15, excel_row_to_restart
  ButtonGroup ButtonPressed
    OkButton 10, 25, 50, 15
    CancelButton 65, 25, 50, 15
  Text 10, 10, 60, 10, "Excel row to start:"
EndDialog

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone
MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr

Do 
    err_msg = ""
    dialog ABAWD_report_dialog
    If buttonpressed = 0 then stopscript								'loops until all errors are resolved
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

If restart_checkbox = 0 then
    
	'Starting the query start time (for the query runtime at the end)
 	query_start_time = timer
	'IF x_number = "" THEN CALL find_variable("User: ", x_number, 7)
	
    'Opening the Excel file
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = True
    Set objWorkbook = objExcel.Workbooks.Add() 
    objExcel.DisplayAlerts = True
	
	'Setting the Excel rows with variables
	ObjExcel.Cells(1, 1).Value = "Worker"
	ObjExcel.Cells(1, 2).Value = "Case Number"
	ObjExcel.Cells(1, 3).Value = "Client Name"
	ObjExcel.Cells(1, 4).Value = "REF #"
	ObjExcel.Cells(1, 5).Value = "PMI #"
	ObjExcel.Cells(1, 6).Value = "SMI #"

    FOR i = 1 to 6		'formatting the cells'
    	objExcel.Cells(1, i).Font.Bold = True		'bold font'
    	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
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
    				
    				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 6  	'Reading case number
    				EMReadScreen client_name, 18, MAXIS_row, 16
    				
                    'if more than one HH member is on the list then non-MEMB 01's don't have a case number listed, this fixes that
    				If trim(MAXIS_case_number) = "" AND trim(client_name) <> "" then 			'if there's a name and no case number
    					EMReadScreen alt_case_number, 8, MAXIS_row - 1, 6				'then it reads the row above
                        MAXIS_case_number = alt_case_number									'restablishes that in this instance, alt case number = case number'    
                    END IF
                    
                    If trim(MAXIS_case_number) = "" and trim(client_name) = "" then exit do			'Exits do if we reach the end
    				
    				'add case/case information to Excel
            		ObjExcel.Cells(excel_row, 1).Value = worker
            		ObjExcel.Cells(excel_row, 2).Value = trim(MAXIS_case_number)
                    ObjExcel.Cells(excel_row, 3).Value = trim(client_name)			
                    
    				excel_row = excel_row + 1	'moving excel row to next row'
    				MAXIS_case_number = ""          'Blanking out variable
    				MAXIS_row = MAXIS_row + 1	'adding one row to search for in MAXIS
    			Loop until MAXIS_row = 19
    			PF8
    			EMReadScreen last_page_check, 21, 24, 2	'Checking for the last page of cases.
    		Loop until last_page_check = "THIS IS THE LAST PAGE"
    	End if
    next
	
	excel_row = 2           're-establishing the row to start checking the members for
Else 
'dialog and dialog DO...Loop	
    Do
    	Do
    		'The dialog is defined in the loop as it can change as buttons are pressed 
    		BeginDialog file_select_dialog, 0, 0, 221, 50, "Select the ABAWD pull cases into Excel file."
    			ButtonGroup ButtonPressed
    			PushButton 175, 10, 40, 15, "Browse...", select_a_file_button
    			OkButton 110, 30, 50, 15
    			CancelButton 165, 30, 50, 15
    			EditBox 5, 10, 165, 15, file_selection_path
    		EndDialog
    		err_msg = ""
    		Dialog file_select_dialog
    		cancel_confirmation
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
    
    do 
    	dialog excel_row_dialog
    	If buttonpressed = 0 then stopscript								'loops until all errors are resolved
    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in
    
    excel_row = excel_row_to_restart
End if 
'Now the script goes back into MFCM and grabs the member # and client name, then cchecks the potentially exempt members for subsidized housing

Do
	MAXIS_case_number  = objExcel.cells(excel_row, 2).Value	're-establishing the case number to use for the case
    client_name        = objExcel.cells(excel_row, 3).Value	're-establishing the client name to use for the case
    If MAXIS_case_number = "" then exit do						'exits do if the case number is ""
	Call navigate_to_MAXIS_screen("REPT", "MFCM")
	
	EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets added to priv case list
	If PRIV_check = "PRIV" then
		ObjExcel.Cells(excel_row, 4).Value = "Privliged case"
		'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
		Do
			back_to_self
			EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
			If SELF_screen_check <> "SELF" then PF3
		LOOP until SELF_screen_check = "SELF"
		EMWriteScreen "________", 18, 43		'clears the case number
		transmit
	Else 
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
			'msgbox row & " for content for member"     
	    Else 
	    	EMWriteScreen "x", 7, 36		'going into the SANC panel to get case info
	    End if 
	
		transmit
	    'For all of the cases that aren't privileged...
        EMReadScreen ERRR_panel_check, 4, 2, 52         'Ensuring that there are no errors on the case. If they are the client inforamiton will not input.
        If ERRR_panel_check = "ERRR" then transmit
	    
	    'Reading and inputing information from the SANC panel
	    EMReadScreen memb_number, 2, 4, 12		'reading member number
	   	ObjExcel.Cells(excel_row, 4).Value = memb_number	
		
		Call navigate_to_MAXIS_screen("STAT", "MEMB")
		Call write_value_and_transmit(memb_number, 20, 76)
		
		EMReadScreen memb_PMI, 10, 4, 46
		EMReadScreen memb_SMI, 10, 5, 46
	        	
        ObjExcel.Cells(excel_row, 5).Value = trim(memb_PMI)		
	    ObjExcel.Cells(excel_row, 6).Value = trim(memb_SMI)
	       
	End if 
	excel_row = excel_row + 1
	STATS_counter = STATS_counter + 1
LOOP UNTIL objExcel.Cells(excel_row, 1).Value = ""	'Loops until there are no more cases in the Excel list

IF priv_case_list <> "" then 
	'Creating the list of privileged cases and adding to the spreadsheet
	excel_row = 2				'establishes the row to start writing the PRIV cases to
	objExcel.cells(1, 8).Value = "PRIV cases"
	
	prived_case_array = split(priv_case_list, "|")
	
	FOR EACH MAXIS_case_number in prived_case_array
		If trim(MAXIS_case_number) <> "" then 
			objExcel.cells(excel_row, 8).value = MAXIS_case_number		'inputs cases into Excel
			excel_row = excel_row + 1								'increases the row
		End if 
	NEXT
End if

'Logging usage stats
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
'------------------------------Post MAXIS coding-----------------------------------------------------------------------------
'Query date/time/runtime info
ObjExcel.Cells(1, 9).Value = "Query date and time:"	'Goes back one, as this is on the next row
objExcel.Cells(1, 9).Font.Bold = TRUE
ObjExcel.Cells(2, 9).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
objExcel.Cells(2, 9).Font.Bold = TRUE
ObjExcel.Cells(3, 9).Value = "Case count:"	'Goes back one, as this is on the next row
objExcel.Cells(3, 9).Font.Bold = TRUE
ObjExcel.Cells(1, 10).Value = now
ObjExcel.Cells(2, 10).Value = timer - query_start_time
ObjExcel.Cells(3, 10).Value = STATS_counter

FOR i = 1 to 10		'formatting the cells'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

script_end_procedure("Success! Please review the list generated.")