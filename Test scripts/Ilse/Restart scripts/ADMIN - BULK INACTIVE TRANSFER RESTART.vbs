'Required for statistical purposes==========================================================================================
name_of_script = "ADMIN - INACTIVE TRANSFER.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 120                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
'END OF stats block=========================================================================================================

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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK=================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
CALL changelog_update("11/04/2022", "Updated to merge BULK/INAC script and transferring cases in one script.", "Ilse Ferris, Hennepin County") '#916
CALL changelog_update("07/01/2022", "Update to ensure run is complete with error handling.", "MiKayla Handley, Hennepin County") '#868'
CALL changelog_update("02/14/2019", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""
Call check_for_MAXIS(False)

transfer_to_worker = "X127CCL" 'setting the worker to the closed basket
excel_row = 2 'default

MAXIS_footer_month = right("0" & DatePart("m", DateAdd("m", -10, date) ), 2) ' resetting the month to current month minus 4
MAXIS_footer_year =  right(DatePart("yyyy", DateAdd("m", -10, date) ), 2)

'The dialog is defined in the loop as it can change as buttons are pressed 
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 115, "Restart BULK Inactive Transfer."
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse...", select_a_file_button
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  EditBox 15, 50, 180, 15, file_selection_path
  Text 20, 20, 235, 25, "This script should be used when BULK INAC Transfer processes needs to be restarted at the point of the transfer portion."
  Text 15, 70, 230, 15, "Select the Excel file that contains your information by selecting the 'Browse' button, and finding the file."
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

Call check_for_MAXIS(False)
excel_row = excel_row_to_restart

Do
    back_to_self 'resetting MAXIS back to self before getting started with the transfers
    transfer_case_action = TRUE 'default

    MAXIS_case_number = objExcel.cells(excel_row, 2).Value
    IF trim(MAXIS_case_number) = "" then exit do '
    previous_worker_number = objExcel.cells(excel_row, 1).Value

        'sorted alphabetically/numerically - excluded x numbers.
	IF  previous_worker_number = "X1274EC" or _
        previous_worker_number = "X127966" or _
        previous_worker_number = "X127AP7" or _
        previous_worker_number = "X127CCL" or _
        previous_worker_number = "X127CSS" or _
        previous_worker_number = "X127EF8" or _
        previous_worker_number = "X127EF9" or _
        previous_worker_number = "X127EH9" or _
        previous_worker_number = "X127EJ1" or _
        previous_worker_number = "X127EM2" or _
        previous_worker_number = "X127EM3" or _
        previous_worker_number = "X127EM4" or _
        previous_worker_number = "X127EN5" or _
        previous_worker_number = "X127EN6" or _
        previous_worker_number = "X127EN8" or _
        previous_worker_number = "X127EN9" or _
        previous_worker_number = "X127EP1" or _
        previous_worker_number = "X127EP2" or _
        previous_worker_number = "X127EP8" or _
        previous_worker_number = "X127EQ6" or _
        previous_worker_number = "X127EQ7" or _
        previous_worker_number = "X127EW4" or _
        previous_worker_number = "X127EW6" or _
        previous_worker_number = "X127EW7" or _
        previous_worker_number = "X127EW8" or _
        previous_worker_number = "X127EX4" or _
        previous_worker_number = "X127EX5" or _
        previous_worker_number = "X127EZ2" or _
        previous_worker_number = "X127F3E" or _
        previous_worker_number = "X127F3F" or _
        previous_worker_number = "X127F3J" or _
        previous_worker_number = "X127F3K" or _
        previous_worker_number = "X127F3N" or _
        previous_worker_number = "X127F3P" or _
        previous_worker_number = "X127F4A" or _
        previous_worker_number = "X127F4B" or _
        previous_worker_number = "X127FE2" or _
        previous_worker_number = "X127FE3" or _
        previous_worker_number = "X127FE6" or _
        previous_worker_number = "X127FF1" or _
        previous_worker_number = "X127FF2" or _
        previous_worker_number = "X127FF4" or _
        previous_worker_number = "X127FF5" or _
        previous_worker_number = "X127FG1" or _
        previous_worker_number = "X127FG2" or _
        previous_worker_number = "X127FG5" or _
        previous_worker_number = "X127FG6" or _
        previous_worker_number = "X127FG7" or _
        previous_worker_number = "X127FG9" or _
        previous_worker_number = "X127FH3" or _
        previous_worker_number = "X127FI1" or _
        previous_worker_number = "X127FI3" or _
        previous_worker_number = "X127FI6" or _
        previous_worker_number = "X127FJ2" or _
        previous_worker_number = "X127GF5" or _
        previous_worker_number = "X127Q95" or _
        previous_worker_number = "X127Y86" THEN
		transfer_case_action  = FALSE
		action_completed = "Excluded"
	Else
		CALL navigate_to_MAXIS_screen_review_PRIV("SPEC", "XFER", is_this_priv)
		IF is_this_priv = TRUE THEN
			transfer_case_action = FALSE
			action_completed = "PRIV"
		ELSE
		    Call write_value_and_transmit("X", 7, 16)              'transfer within county option
	        PF9                                                    'putting the transfer in edit mode
            EMreadscreen second_servicing_worker, 7, 18, 74
	        IF second_servicing_worker <> "_______" THEN CALL clear_line_of_text(18, 74)
            Call write_value_and_transmit(transfer_to_worker, 18, 61)           'entering the worker information
            EMReadScreen servicing_worker, 7, 24, 30
            If servicing_worker <> transfer_to_worker THEN     'if it is not the transfer_to_worker - the transfer failed.
				EMReadScreen MISC_error_check,  74, 24, 02
                transfer_case_action = FALSE
				action_completed = trim(MISC_error_check)
	        Else
				action_completed = "Successful transfer."
            End if
		END IF
	END IF
	'Export data to Excel
	ObjExcel.Cells(excel_row, 6).Value = trim(transfer_case_action)
	objExcel.cells(excel_row, 7).Value = trim(action_completed)
	excel_row = excel_row + 1	'increments the excel row so we don't overwrite our data
Loop

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)

'Query date/time/runtime info
objExcel.Cells(1, 8).Font.Bold = TRUE
objExcel.Cells(2, 8).Font.Bold = TRUE
ObjExcel.Cells(1, 8).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, 9).Value = now
ObjExcel.Cells(2, 8).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, 9).Value = timer - query_start_time

FOR i = 1 to 9
	objExcel.Columns(i).AutoFit()				'sizing the columns
NEXT

objWorkbook.Save()  'keeping open to review

script_end_procedure("Success! The list is complete. Please review the cases that appear to be in error.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------11/04/2022
'--Tab orders reviewed & confirmed----------------------------------------------11/04/2022
'--Mandatory fields all present & Reviewed--------------------------------------11/04/2022------------------N/A
'--All variables in dialog match mandatory fields-------------------------------11/04/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------11/04/2022------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------11/04/2022------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------11/04/2022------------------N/A
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used-11/04/2022------------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------11/04/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------11/04/2022------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------11/04/2022
'--Out-of-County handling reviewed----------------------------------------------11/04/2022------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------11/04/2022
'--BULK - review output of statistics and run time/count (if applicable)--------11/04/2022
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---11/04/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------11/04/2022
'--Incrementors reviewed (if necessary)-----------------------------------------11/04/2022
'--Denomination reviewed -------------------------------------------------------11/04/2022
'--Script name reviewed---------------------------------------------------------11/04/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------11/04/2022

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------11/04/2022
'--comment Code-----------------------------------------------------------------11/04/2022
'--Update Changelog for release/update------------------------------------------11/04/2022
'--Remove testing message boxes-------------------------------------------------11/04/2022
'--Remove testing code/unnecessary code-----------------------------------------11/04/2022
'--Review/update SharePoint instructions----------------------------------------11/04/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------11/04/2022
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------11/04/2022
'--Complete misc. documentation (if applicable)---------------------------------11/04/2022
'--Update project team/issue contact (if applicable)----------------------------11/04/2022
