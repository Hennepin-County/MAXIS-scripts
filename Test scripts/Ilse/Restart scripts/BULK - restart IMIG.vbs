'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - restart IMIG.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 60                      'manual run time in seconds
STATS_denomination = "M"       			   'M is for each CASE
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
call changelog_update("08/07/2019", "Updated coding to read citizenship verif code at new location due to MEMI panel changes associated with New Spouse Income Policy.", "Ilse Ferris, Hennepin County")
call changelog_update("09/13/2017", "Updated to remove blank FSET/ABAWD codes for members that do not have a WREG panel", "Ilse Ferris, Hennepin County")
call changelog_update("09/08/2017", "Updated to include FSET codes in addition to ABAWD codes.", "Ilse Ferris, Hennepin County")
call changelog_update("07/31/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone
MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr

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
	dialog Dialog1 
	cancel_without_confirmation
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

excel_row = excel_row_to_restart

Do 
	MAXIS_case_number = ObjExcel.Cells(excel_row, 2).Value
	MAXIS_case_number = trim(MAXIS_case_number)
    
	CALL navigate_to_MAXIS_screen("STAT", "MEMB")
	EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets added to priv case list
	If PRIV_check = "PRIV" then
		objExcel.Cells(excel_row, 5).Value = "PRIV case"
	Else 
		EMReadScreen access_check, 13, 24, 2
		If access_check = "ACCESS DENIED" then 
			objExcel.Cells(excel_row, 5).Value = "Unable to access case."
			PF10
			PF3
		Else 
		
	        DO
	        	EMReadScreen needs_trans, 1, 14, 68
	        	EMReadScreen ref_num, 2, 4, 33
	        	EMReadScreen alien_id, 10, 15, 68	
	        	alien_id = replace(alien_id, "_", "")
	        	
	        	objExcel.Cells(excel_row, 6).Value = objExcel.Cells(excel_row, 6).Value & needs_trans & ";"
	        	 
	        	EMWriteScreen "MEMI", 20, 71
	        	EMWriteScreen ref_num, 20, 76
	        	transmit

	        	EMReadScreen citizen_yn, 1, 11, 49
	        	IF citizen_yn = "Y" THEN objExcel.Cells(excel_row, 5).Value = objExcel.Cells(excel_row, 5).Value & ref_num & ";"
	        		
	        	EMWriteScreen "MEMB", 20, 71
	        	EMWriteScreen ref_num, 20, 76
	        	transmit
	        	
	        	transmit
	        	EMReadScreen error_message, 20, 24, 2
	        	error_message = trim(error_message)
	        LOOP UNTIL error_message <> ""		
	        
	        IF InStr(objExcel.Cells(excel_row, 6).Value, "Y") <> 0 THEN 
	        	objExcel.Cells(excel_row, 6).Value = ""
	        ELSE
	        	objExcel.Cells(excel_row, 6).Value = "N"
	        END IF
		End if 
	End if 
	
	excel_row = excel_row + 1
Loop until ObjExcel.Cells(excel_row, 2).Value = ""

STATS_counter = STATS_counter - 1 'since we start with 1
script_end_procedure("Success! Please review your ABAWD list.")