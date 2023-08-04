'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - GET CASE STATUS.vbs"
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
call changelog_update("05/29/2019", "Updated script to work with new BOBI query.", "Ilse Ferris, Hennepin County")
call changelog_update("01/31/2019", "Added functionality to change Defer FSET funds field if coded incorrectly on STAT/WREG.", "Ilse Ferris, Hennepin County")
call changelog_update("05/23/2018", "Added code to write in client name if presenting as a PRIV case on initial spreadsheet.", "Ilse Ferris, Hennepin County")
call changelog_update("03/30/2018", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Custom function for this script----------------------------------------------------------------------------------------------------
FUNCTION get_case_status
	back_to_self
	EMWriteScreen MAXIS_case_number, 18, 43

	Call navigate_to_MAXIS_screen("CASE", "CURR")
	EMReadScreen CURR_panel_check, 4, 2, 55
	If CURR_panel_check <> "CURR" then ObjExcel.Cells(excel_row, 2).Value = ""

	EMReadScreen case_status, 8, 8, 9
	case_status = trim(case_status)
	ObjExcel.Cells(excel_row, 2).Value = case_status
    excel_row = excel_row + 1

    'supporting if there are multiple DAIL messages for a case
    Do
        next_case_number = objExcel.cells(excel_row, 3).value
        If trim(next_case_number) = trim(MAXIS_case_number) then
            ObjExcel.Cells(excel_row, 2).Value = case_status
            excel_row = excel_row + 1
        Else
            exit do
        End if
    Loop until trim(next_case_number) <> trim(MAXIS_case_number)

	MAXIS_case_number = ""
	'using new variable count to calculate percentages
	IF case_status = "ACTIVE" then active_status = active_status + 1
	IF case_status = "APP OPEN" then active_status = active_status + 1

	IF case_status = "APP CLOS" then inactive_status = inactive_status + 1
	IF case_status = "INACTIVE" then inactive_status = inactive_status + 1

	If case_status = "CAF2 PEN" then pending_status = pending_status + 1
	If case_status = "CAF1 PEN" then pending_status = pending_status + 1

	IF case_status = "REIN" then rein_status = rein_status + 1
	STATS_counter = STATS_counter + 1
END FUNCTION

'----------------------------------------------------------------------------------------------------The script
'CONNECTS TO BlueZone
EMConnect ""
Call check_for_MAXIS(False)

'dialog and dialog DO...Loop
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 301, 100, "BULK - GET CASE STATUS"
  ButtonGroup ButtonPressed
    PushButton 15, 40, 60, 15, "Browse...", select_a_file_button
  EditBox 80, 40, 205, 15, file_selection_path
  ButtonGroup ButtonPressed
    OkButton 190, 80, 50, 15
    CancelButton 245, 80, 50, 15
  Text 40, 60, 230, 10, "This script should be used when a list of cases needs a case status."
  Text 15, 15, 275, 20, "Select the Excel file that contains your information by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 285, 70, "Using this script:"
EndDialog


'dialog and dialog DO...Loop
Do
    Do
        err_msg = ""
        dialog Dialog1
        cancel_without_confirmation
        If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
        If trim(file_selection_path) = "" then err_msg = err_msg & vbcr & "* Select a file to continue."
        If err_msg <> "" Then MsgBox err_msg
    Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Opening today's list
Call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 126, 50, "Select the excel row to start"
  EditBox 75, 5, 40, 15, excel_row_to_restart
  ButtonGroup ButtonPressed
    OkButton 10, 25, 50, 15
    CancelButton 65, 25, 50, 15
  Text 10, 10, 60, 10, "Excel row to start:"
EndDialog

DO
    dialog Dialog1
    cancel_without_confirmation
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

excel_row = excel_row_to_restart

back_to_self
EMWriteScreen CM_mo, 20, 43
EMWriteScreen CM_yr, 20, 46
transmit

Do
	'Grabs the case number
	MAXIS_case_number = objExcel.cells(excel_row, 3).value
	If trim(MAXIS_case_number) = "" then exit do
	Call get_case_status
LOOP UNTIL objExcel.Cells(excel_row, 3).value = ""
STATS_counter = STATS_counter - 1 'removes one from the count since 1 is counted at the beginning

script_end_procedure("All done.")
