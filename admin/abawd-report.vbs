'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - ABAWD report.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 60                      'manual run time in seconds
STATS_denomination = "M"       			   'M is for each CASE
'END OF stats block==============================================================================================

''LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
'IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
'	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
'		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
'			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
'		Else											'Everyone else should use the release branch.
'			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
'		End if
'		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
'		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
'		req.send													'Sends request
'		IF req.Status = 200 THEN									'200 means great success
'			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
'			Execute req.responseText								'Executes the script code
'		ELSE														'Error message
'			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
'                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
'                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
'                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
'            StopScript
'		END IF
'	ELSE
'		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
'		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
'		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
'		text_from_the_other_script = fso_command.ReadAll
'		fso_command.Close
'		Execute text_from_the_other_script
'	END IF
'END IF
''END FUNCTIONS LIBRARY BLOCK================================================================================================
'
'LOADING FUNC LIB
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\MASTER FUNCTIONS LIBRARY.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("09/13/2017", "Updated to remove blank FSET/ABAWD codes for members that do not have a WREG panel", "Ilse Ferris, Hennepin County")
call changelog_update("09/08/2017", "Updated to include FSET codes in addition to ABAWD codes.", "Ilse Ferris, Hennepin County")
call changelog_update("07/31/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

Function HCRE_panel_bypass() 
	'handling for cases that do not have a completed HCRE panel
	PF3		'exits PROG to prommpt HCRE if HCRE insn't complete
	Do
		EMReadscreen HCRE_panel_check, 4, 2, 50
		If HCRE_panel_check = "HCRE" then
			PF10	'exists edit mode in cases where HCRE isn't complete for a member
			PF3
		END IF
	Loop until HCRE_panel_check <> "HCRE"
End Function

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
    
    'Setting the first 3 col as worker, case number, and name
    ObjExcel.Cells(1, 1).Value = "X Number"
    ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
    ObjExcel.Cells(1, 3).Value = "NAME"
	ObjExcel.Cells(1, 4).Value = "NEXT REVW DATE"
	ObjExcel.Cells(1, 5).Value = "ABAWD CODES"
	
	FOR i = 1 to 5		'formatting the cells'
		objExcel.Cells(1, i).Font.Bold = True		'bold font'
		ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
		objExcel.Columns(i).AutoFit()				'sizing the columns'
	NEXT
	
	excel_row = 2
	
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
		IF worker <> "" THEN EMWriteScreen worker, 21, 13
		transmit
	
		'Grabbing each case number on screen
		Do
			row = 7
			EMReadScreen last_page_check, 21, 24, 2
			Do
				EMReadScreen MAXIS_case_number, 8, row, 12
				If trim(MAXIS_case_number) = "" then exit do
				EMReadScreen SNAP_status, 1, row, 61
				If SNAP_status = "A" then 
					add_to_excel = true 
				Else
					EMReadScreen cash_status, 9, row, 51
					cash_status = trim(cash_status)
					If instr(cash_status, "MF A") then 
						add_to_excel = true
					Else 
						add_to_excel = False 
					End if 
				End if 
				
				'msgbox MAXIS_case_number & vbcr & add_to_excel
				
				If add_to_excel = True then 
					EMReadScreen client_name, 21, row, 21
					EMReadScreen next_REVW_date, 8, row, 42
					ObjExcel.Cells(excel_row, 1).Value = worker
					ObjExcel.Cells(excel_row, 2).Value = MAXIS_case_number
					ObjExcel.Cells(excel_row, 3).Value = client_name
					ObjExcel.Cells(excel_row, 4).Value = replace(next_REVW_date, " ", "/")
					excel_row = excel_row + 1
					Stats_counter = stats_counter + 1
				End if 
				row = row + 1
			Loop until row = 19
			PF8
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	next
	
	'Resetting excel_row variable, now we need to start looking people up
	excel_row = 2 
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
ENd if 

Do 
	MAXIS_case_number = ObjExcel.Cells(excel_row, 2).Value
	MAXIS_case_number = trim(MAXIS_case_number)
    
	'Now pulling ABAWD info
    ABAWD_status = "" 		'clearing variable
	ABAWD_info = ""
    eats_group_members = ""		'clearing
    eats_row = 13			'clearing variable
	'msgbox MAXIS_case_number
    	
    call navigate_to_MAXIS_screen("STAT", "EATS")
	EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets added to priv case list
	If PRIV_check = "PRIV" then
		objExcel.Cells(excel_row, 5).Value = "PRIV case"
	Else 
	
        EMReadScreen all_eat_together, 1, 4, 72
	    
	    'Handling for single HH member
        IF all_eat_together = "_" THEN		
	    	call navigate_to_MAXIS_screen("STAT", "WREG")
	    	EMReadScreen FSET_code, 2, 8, 50
	    	EMReadScreen ABAWD_status_code, 2, 13, 50
	    	ABAWD_info = FSET_code & "/" & ABAWD_status_code
	    	ABAWD_status = "01: " & ABAWD_info 
	    	
	    	ObjExcel.Cells(excel_row, 5).Value = ABAWD_status
	    	STATS_counter = STATS_counter + 1
	    	'more than one HH member
        ELSEIF all_eat_together = "Y" THEN 
        	eats_row = 5
        	DO
        		EMReadScreen eats_person, 2, eats_row, 3
        		eats_person = trim(eats_person)
        		IF eats_person <> "" THEN 
        			eats_group_members = eats_group_members & eats_person & " "
        			eats_row = eats_row + 1
        		END IF
        	LOOP UNTIL eats_person = "" or eats_row = 18
        ELSEIF all_eat_together = "N" THEN
        	eats_row = 13
        	DO
        		EMReadScreen eats_group, 38, eats_row, 39
        		find_memb01 = InStr(eats_group, "01")
        		IF find_memb01 = 0 THEN 
	    			eats_row = eats_row + 1
	    		else 
	    			exit do 
	    		End if 
        	LOOP UNTIL find_memb01 <> 0 OR eats_row = 18
        	IF eats_row <> 18 THEN 
        		eats_col = 39
        		DO
        			EMReadScreen eats_group, 2, eats_row, eats_col
        			IF eats_group <> "__" THEN 
        				eats_group_members = eats_group_members & eats_group & " "
        				eats_col = eats_col + 4
        			END IF
        		LOOP UNTIL eats_group = "__"
	    	END IF 
	    End if 
	    		
	    IF eats_row <> 18 then 
	    	eats_group_members = trim(eats_group_members)
	    	eats_group_members = split(eats_group_members)

	    	call navigate_to_MAXIS_screen("STAT", "WREG")

	    	FOR EACH person IN eats_group_members
	    		STATS_counter = STATS_counter + 1
	    		EMWriteScreen person, 20, 76
	    		transmit
	    		
	    		EMReadScreen FSET_code, 2, 8, 50
	    		EMReadScreen ABAWD_status_code, 2, 13, 50
	    		ABAWD_info = FSET_code & "/" & ABAWD_status_code
	    		If ABAWD_info <> "__/__" then ABAWD_status = ABAWD_status & person & ": " & ABAWD_info & ","
	    	NEXT

	    	ObjExcel.Cells(excel_row, 5).Value = ABAWD_status	
        ELSE 
        	objExcel.Cells(excel_row, 5).Value = "CHECK MANUALLY"
	    	STATS_counter = STATS_counter + 1
        END IF
	End if
	excel_row = excel_row + 1
Loop until ObjExcel.Cells(excel_row, 2).Value = ""

STATS_counter = STATS_counter - 1 'since we start with 1
script_end_procedure("Success! Please review your ABAWD list.")