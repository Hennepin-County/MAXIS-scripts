'Required for statistical purposes==========================================================================================
name_of_script = "BULK - DISPLAY LIMIT TRANSFER.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 229                	'manual run time in seconds
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
call changelog_update("02/14/2019", "Initial version.", "MiKayla Handley")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
EMConnect ""
'------------------------------------------------------------------------THE SCRIPT
Dialog1 = ""
BeginDialog dialog1, 0, 0, 316, 65, "Select the source file"
  EditBox 5, 25, 260, 15, file_selection_path
  ButtonGroup ButtonPressed
  PushButton 270, 25, 40, 15, "Browse:", select_a_file_button
  OkButton 205, 45, 50, 15
  CancelButton 260, 45, 50, 15
  Text 5, 5, 295, 15, "Click the BROWSE button and select the INAC report for today. Once selected, click 'OK'. There will be no additional input needed until the script run is complete."
EndDialog
Do
'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    'Show initial dialog
    Do
		Dialog dialog1
		cancel_without_confirmation
		If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
	Loop until ButtonPressed = OK and file_selection_path <> ""
	If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

ObjExcel.Cells(1, 1).Value = "WORKER"
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
ObjExcel.Cells(1, 3).Value = "CASE NAME"
ObjExcel.Cells(1, 4).Value = "REVW"

ObjExcel.Cells(1, 8).Value = "TRANSFERED"
ObjExcel.Cells(1, 9).Value = "CONFRIM"

FOR i = 1 to 9		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	'ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'This bit freezes the top row of the Excel sheet for better use ability when there is a lot of information
ObjExcel.ActiveSheet.Range("A2").Select
objExcel.ActiveWindow.FreezePanes = True


'Now the script adds all the clients on the excel list into an array
excel_row = 2 're-establishing the row to start checking the members for
'entry_record = 0
transfer_case_action = TRUE
Do
    previous_worker_number = objExcel.cells(excel_row, 1).Value          're-establishing the worker number for functions to use
    If previous_worker_number = "" then exit do
    previous_worker_number = trim(previous_worker_number)
    MAXIS_case_number 	 = objExcel.cells(excel_row, 2).Value          're-establishing the case numbers for functions to use
    If MAXIS_case_number = "" then exit do
    MAXIS_case_number	 = trim(MAXIS_case_number)
	'msgbox previous_worker_number & " / " & transfer_case_action
	CALL navigate_to_MAXIS_screen ("SPEC", "XFER")
	EMWriteScreen MAXIS_case_number, 18, 43
	TRANSMIT
	EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets added to priv case list
	If PRIV_check = "PRIV" then
		transfer_case_action = FALSE
		action_completed = FALSE	'row gets deleted since it will get added to the priv case list at end of script
		IF excel_row = 2 then
			excel_row = excel_row
		Else
			excel_row = excel_row - 1
		End if
		'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
		Do
			back_to_self
			EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
			If SELF_screen_check <> "SELF" then PF3
		LOOP until SELF_screen_check = "SELF"
		EMWriteScreen "________", 18, 43		'clears the case number
		transmit
		msgbox PRIV                                                           'Loops until there are no more cases in the Excel list
	ELSE
		'CALL navigate_to_MAXIS_screen ("SPEC", "XFER")
    	Call write_value_and_transmit("x", 7, 16) 'This should have us in SPEC/XWKR'
    	EMReadScreen panel_check, 4, 2, 55
    	IF panel_check <> "XWKR" THEN MsgBox panel_check
    	EMReadScreen prev_worker, 7, 18, 28
    	'MsgBox prev_worker & " / " & previous_worker_number & " / " & transfer_case_action
    	'If prev_worker = previous_worker_number THEN transfer_case_action = FALSE
    	PF9
    	'MsgBox "writing"
    	EMWriteScreen servicing_worker, 18, 61
    	CALL clear_line_of_text(18, 74)

    	TRANSMIT

    	EMReadScreen worker_check, 9, 24, 2
    	IF worker_check = "SERVICING" or worker_check = "LAST" THEN
      		action_completed = False
       		PF10
    	END IF
       	EMReadScreen transfer_confirmation, 16, 24, 2
       	IF transfer_confirmation = "CASE XFER'D FROM" then
       		action_completed = True
       	Else
       		action_completed = False
       	End if
       	PF3
        'excel_row = excel_row + 1	'increments the excel row so we don't overwrite our data
    ELSE
        transfer_case_action = FALSE
        action_completed = FALSE
	    excel_row = excel_row + 1	'increments the excel row so we don't overwrite our data
        'END IF
	END IF
	'Export data to Excel
		ObjExcel.Cells(excel_row, 8).Value = trim(transfer_case_action)
		objExcel.cells(excel_row, 9).Value = trim(action_completed)
		excel_row = excel_row + 1	'increments the excel row so we don't overwrite our data
LOOP UNTIL previous_worker_number = ""

script_end_procedure("Success! The list is complete. Please review the cases that appear to be in error.")
