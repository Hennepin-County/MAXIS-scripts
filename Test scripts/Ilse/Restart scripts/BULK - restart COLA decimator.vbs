'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - COLA DECIMATOR.vbs"
start_time = timer
STATS_counter = 1                       'sets the stats counter at one
STATS_manualtime = 20
STATS_denomination = "C"       			'C is for each CASE
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
call changelog_update("06/11/2018", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display

'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""

'The dialog is defined in the loop as it can change as buttons are pressed 
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 115, "Restart COLA Decimator at CASE/NOTE."
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse...", select_a_file_button
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  EditBox 15, 50, 180, 15, file_selection_path
  Text 20, 20, 235, 25, "This script should be used when a COLA Decimator list needs to be restared at the point of the Case noting portion."
  Text 15, 70, 230, 15, "Select the Excel file that contains your inforamtion by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 250, 85, "Using this script:"
EndDialog
'dialog and dialog DO...Loop	
Do
    'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    'Show initial dialog
    Do
    	Dialog Dialog1 
    	cancel_without_confirmation
    	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
    Loop until ButtonPressed = OK and file_selection_path <> ""
    If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Dialog1 = ""
'Select Excel row dialog
BeginDialog Dialog1, 0, 0, 126, 50, "Select the excel row to restart"
  EditBox 75, 5, 40, 15, excel_row_to_restart
  ButtonGroup ButtonPressed
    OkButton 10, 25, 50, 15
    CancelButton 65, 25, 50, 15
  Text 10, 10, 60, 10, "Excel row to start:"
EndDialog

Do 
	dialog Dialog1 
	cancel_without_confirmation
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

excel_row = excel_row_to_restart

Do 
    MAXIS_case_number = ObjExcel.Cells(excel_row, 2).Value
    MAXIS_case_number = trim(MAXIS_case_number)

    dail_msg = ObjExcel.Cells(excel_row, 5).Value
    dail_msg = trim(dail_msg)
    'Cleaning up the DAIL messages for the case note 
    If right(dail_msg, 9) = "-SEE PF12" THEN dail_msg = left(dail_msg, len(dail_msg) - 9)
    If right(dail_msg, 1) = "*" THEN dail_msg = left(dail_msg, len(dail_msg) - 1)
    dail_msg = trim(dail_msg)
    
    Call navigate_to_MAXIS_screen("CASE", "NOTE")
    EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets added to priv case list
    
    If PRIV_check = "PRIV" then 
        objExcel.Cells(excel_row, 6).Value = "PRIV, unable to case note."
        'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
    	Do
    		back_to_self
    		EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
    		If SELF_screen_check <> "SELF" then PF3
    	LOOP until SELF_screen_check = "SELF"
    	EMWriteScreen "________", 18, 43		'clears the MAXIS case number
    	transmit
    Else
        EmReadscreen county_check, 2, 21, 16
        If county_check <> "27" then 
            objExcel.Cells(excel_row, 6).Value = "Out of county case."
        Else 
            Call start_a_blank_CASE_NOTE
            CALL write_variable_in_case_note(dail_msg)
            PF3 ' save message
            objExcel.Cells(excel_row, 6).Value = "Case note created."
        End if 
    End If 
    excel_row = excel_row + 1     
Loop until ObjExcel.Cells(excel_row, 2).Value = ""    

STATS_counter = STATS_counter - 1
'Enters info about runtime for the benefit of folks using the script
objExcel.Cells(2, 7).Value = "Number of DAILs deleted:"
objExcel.Cells(3, 7).Value = "Average time to find/select/copy/paste one line (in seconds):"
objExcel.Cells(4, 7).Value = "Estimated manual processing time (lines x average):"
objExcel.Cells(5, 7).Value = "Script run time (in seconds):"
objExcel.Cells(6, 7).Value = "Estimated time savings by using script (in minutes):"
objExcel.Cells(7, 7).Value = "Number of " & dail_to_decimate & " messages reviewed"
objExcel.Columns(7).Font.Bold = true
objExcel.Cells(2, 8).Value = deleted_dails
objExcel.Cells(3, 8).Value = STATS_manualtime
objExcel.Cells(4, 8).Value = STATS_counter * STATS_manualtime
objExcel.Cells(5, 8).Value = timer - start_time
objExcel.Cells(6, 8).Value = ((STATS_counter * STATS_manualtime) - (timer - start_time)) / 60
objExcel.Cells(7, 8).Value = STATS_counter

'Formatting the column width.
FOR i = 1 to 8
	objExcel.Columns(i).AutoFit()
NEXT

script_end_procedure("Success! Please review the list created for accuracy.")