'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - GET HSS NUMBER.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 20
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================

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

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
'The dialog is defined in the loop as it can change as buttons are pressed
BeginDialog Dialog1, 0, 0, 221, 50, "Select the case list source file"
    ButtonGroup ButtonPressed
    PushButton 175, 10, 40, 15, "Browse...", select_a_file_button
    OkButton 110, 30, 50, 15
    CancelButton 165, 30, 50, 15
    EditBox 5, 10, 165, 15, file_selection_path
EndDialog

'dialog and dialog DO...Loop
Do
    'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    'Show initial dialog
    Do
        err_msg = ""
    	Dialog Dialog1
    	If ButtonPressed = cancel then stopscript
    	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
        If file_selection_path = "" then err_msg = err_msg & vbNewLine & "Use the Browse Button to select the file that has your client data"
        If err_msg <> "" Then MsgBox err_msg
    Loop until ButtonPressed = OK and file_selection_path <> ""
    If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in


excel_row = 2

back_to_self
EMwritescreen "________", 18, 43
EMWriteScreen CM_mo, 20, 43
EMWriteScreen CM_yr, 20, 46

worker_number_col = 2
hss_col = 1
case_num_col = 3
DO
    'Grabs the case number
	worker_number = objExcel.cells(excel_row, worker_number_col).value
    If worker_number = "" then exit do
	back_to_self
    Call navigate_to_MAXIS_screen("REPT", "USER")
    EMReadScreen CURR_panel_check, 4, 2, 47
	If CURR_panel_check <> "USER" then ObjExcel.Cells(excel_row, worker_number_col).Value = ""

	EmWriteScreen worker_number, 21, 12
	TRANSMIT
	'MsgBox "where 1"
	EmReadscreen worker_number_check, 7, 7, 5
	if worker_number = worker_number_check then
		EmWriteScreen "X", 7, 3
	    TRANSMIT
	    'MsgBox "cursor"
	    EMReadScreen HSS, 7, 14, 61
	    'MsgBox "where"
	    ObjExcel.Cells(excel_row, hss_col).Value = HSS
	    TRANSMIT
        MAXIS_case_number = ""
        basket = ""
        HSS = ""
        excel_row = excel_row + 1
        STATS_counter = STATS_counter + 1
	else
		MsgBox "error"
	END IF
LOOP UNTIL objExcel.Cells(excel_row, worker_number_col).value = ""	'looping until the list of cases to check for recert is complete

STATS_counter = STATS_counter - 1 'removes one from the count since 1 is counted at the beginning (because counting :p)
script_end_procedure("Success! The Excel file now has been updated. Please review the blank case statuses that remain.")
