'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - ADD NEXT REVW DATE.vbs"
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
call changelog_update("02/08/2018", "Initial version.", "Ilse Ferris, Hennepin County")

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
        BeginDialog Dialog1, 0, 0, 221, 50, "Select the GRH file to update."
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
	MAXIS_case_number= objExcel.cells(excel_row, 2).Value	're-establishing the case number to use for the case
    If trim(MAXIS_case_number) = "" then exit do
	
	'This Do...loop gets back to SELF
	back_to_self
	call navigate_to_MAXIS_screen("rept", "ACTV")
    EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets added to priv case list
	If PRIV_check = "PRIV" then
		ObjExcel.Cells(excel_row, 4).Value = "PRIV cases"
		'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
		Do
			back_to_self
			EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
			If SELF_screen_check <> "SELF" then PF3
		LOOP until SELF_screen_check = "SELF"
		EMWriteScreen "________", 18, 43		'clears the case number
		transmit
    Else 
	    EMReadScreen next_revw_date, 8, 7, 42
        ObjExcel.Cells(excel_row, 4).Value = replace(next_revw_date, " ", "/")
    End if 
    excel_row = excel_row + 1 'setting up the script to check the next row.
LOOP UNTIL objExcel.Cells(excel_row, 2).Value = ""	'Loops until there are no more cases in the Excel list
        
'formatting the cells
FOR i = 1 to 9
	objExcel.Columns(i).AutoFit()				'sizing the columns
NEXT

STATS_counter = STATS_counter - 1 'since we start with 1
script_end_procedure("Success! Please review your ABAWD list.")