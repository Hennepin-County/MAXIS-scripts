'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - GET CASE CURR STATUSES.vbs"
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
call changelog_update("08/18/2021", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------The script
'CONNECTS TO BlueZone
EMConnect ""
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr

case_number_col = 2
file_selection_path = ""

'dialog and dialog DO...Loop
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 301, 100, "ADMIN - GET CASE CURR STATUSES.vbs"
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

Do
    'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    'Show initial dialog
    Do
    	Dialog Dialog1
    	If ButtonPressed = cancel then stopscript
    	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
    Loop until ButtonPressed = OK and file_selection_path <> ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'

'Dialog1 = ""
'BeginDialog Dialog1, 0, 0, 126, 50, "Select the excel row to start"
'  EditBox 75, 5, 40, 15, excel_row_to_restart
'  ButtonGroup ButtonPressed
'    OkButton 10, 25, 50, 15
'    CancelButton 65, 25, 50, 15
'  Text 10, 10, 60, 10, "Excel row to start:"
'EndDialog
'
'DO
'    dialog Dialog1
'    If buttonpressed = 0 then stopscript								'loops until all errors are resolved
'    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
'LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
'
'excel_row = excel_row_to_restart
excel_row = 2

back_to_self
EMWriteScreen CM_mo, 20, 43
EMWriteScreen CM_yr, 20, 46
transmit

Do
	'Grabs the case number
	MAXIS_case_number = objExcel.cells(excel_row, case_number_col).value
	If MAXIS_case_number = "" then exit do
	Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
    ObjExcel.Cells(excel_row,  6).Value = case_active
    ObjExcel.Cells(excel_row,  7).Value = case_pending
    ObjExcel.Cells(excel_row,  8).Value = case_rein
    ObjExcel.Cells(excel_row,  9).Value = family_cash_case
    ObjExcel.Cells(excel_row, 10).Value = mfip_case
    ObjExcel.Cells(excel_row, 11).Value = dwp_case
    ObjExcel.Cells(excel_row, 12).Value = adult_cash_case
    ObjExcel.Cells(excel_row, 13).Value = ga_case
    ObjExcel.Cells(excel_row, 14).Value = msa_case
    ObjExcel.Cells(excel_row, 15).Value = grh_case
    ObjExcel.Cells(excel_row, 16).Value = snap_case
    ObjExcel.Cells(excel_row, 17).Value = ma_case
    ObjExcel.Cells(excel_row, 18).Value = msp_case
    ObjExcel.Cells(excel_row, 19).Value = unknown_cash_pending
    ObjExcel.Cells(excel_row, 20).Value = unknown_hc_pending
    ObjExcel.Cells(excel_row, 21).Value = ga_status
    ObjExcel.Cells(excel_row, 22).Value = msa_status
    ObjExcel.Cells(excel_row, 23).Value = mfip_status
    ObjExcel.Cells(excel_row, 24).Value = dwp_status
    ObjExcel.Cells(excel_row, 25).Value = grh_status
    ObjExcel.Cells(excel_row, 26).Value = snap_status
    ObjExcel.Cells(excel_row, 27).Value = ma_status
    ObjExcel.Cells(excel_row, 28).Value = msp_status

    STATS_counter = STATS_counter + 1
    excel_row = excel_row + 1
LOOP UNTIL objExcel.Cells(excel_row, case_number_col).value = ""
STATS_counter = STATS_counter - 1 'removes one from the count since 1 is counted at the beginning

script_end_procedure("All done.")
