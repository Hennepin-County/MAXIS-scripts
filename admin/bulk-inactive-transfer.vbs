'Required for statistical purposes==========================================================================================
name_of_script = "BULK - INACTIVE TRANSFER.vbs"
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
'------------------------------------------------------------------------THE SCRIPT
EMConnect ""
BeginDialog dialog1, 0, 0, 316, 65, "Select the source file"
  EditBox 5, 25, 260, 15, file_selection_path
  ButtonGroup ButtonPressed
  PushButton 270, 25, 40, 15, "Browse...", select_a_file_button
  OkButton 205, 45, 50, 15
  CancelButton 260, 45, 50, 15
  Text 5, 5, 295, 15, "Click the BROWSE button and select the INAC report for today. Once selected, click 'OK'. There will be no additional input needed until the script run is complete."
EndDialog
Do
'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    'Show initial dialog
    Do
		Dialog dialog1
		If ButtonPressed = cancel then stopscript
		If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
	Loop until ButtonPressed = OK and file_selection_path <> ""
	If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

ObjExcel.Cells(1, 1).Value = "WORKER"
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
ObjExcel.Cells(1, 3).Value = "CASE NAME"
ObjExcel.Cells(1, 4).Value = "APPL DATE"
ObjExcel.Cells(1, 5).Value = "INAC DATE"
ObjExcel.Cells(1, 6).Value = "TRANSFERED"
ObjExcel.Cells(1, 7).Value = "CONFRIM"

FOR i = 1 to 8		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	'ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'This bit freezes the top row of the Excel sheet for better use ability when there is a lot of information
ObjExcel.ActiveSheet.Range("A2").Select
objExcel.ActiveWindow.FreezePanes = True

'Sets up the array to store all the information for each client'
Dim INAC_array()
ReDim INAC_array (7, 0)

'Sets constants for the array to make the script easier to read (and easier to code)'
Const worker_number    		= 1			'Each of the case numbers will be stored at this position'
Const case_number      		= 2
Const case_member_name		= 3
Const appl_date			  	= 4
Const inac_date				= 5
Const trans_status	 	  	= 6
Const trans_conf	    	= 7

'Now the script adds all the clients on the excel list into an array
excel_row = 2 're-establishing the row to start checking the members for
'entry_record = 0
transfer_case_action = TRUE

Do                                                            'Loops until there are no more cases in the Excel list
	previous_worker_number = objExcel.cells(excel_row, 1).Value          're-establishing the worker number for functions to use
	If previous_worker_number = "" then exit do
	previous_worker_number = trim(previous_worker_number)

	IF previous_worker_number = "X1274EC" THEN transfer_case_action = FALSE
	'msgbox "First: " & previous_worker_number & " " & transfer_case_action

	MAXIS_case_number 	 = objExcel.cells(excel_row, 2).Value          're-establishing the case numbers for functions to use
	If MAXIS_case_number = "" then exit do
	MAXIS_case_number	 = trim(MAXIS_case_number)

	case_name  			= objExcel.cells(excel_row,  3).value	'(col I) establishes
	case_name	 		= trim(case_name)

	application_date 	= objExcel.cells(excel_row, 4).value	'(col K) establishes claim number from MAXIS
	application_date 	= trim(application_date)

    inactive_date 		= objExcel.cells(excel_row, 5).value	'(col P) establishes
	inactive_date   	= trim(inactive_date)

	transfer_case 		= objExcel.cells(excel_row, 6).value	'(col Q) establishes grant amount for each case
	transfer_case		= trim(transfer_case)

	transfer_confirmed 	= objExcel.cells(excel_row, 7).value	'(col R) establishes
	transfer_confirmed		= trim(transfer_confirmed)

	'Adding client information to the array'
	'ReDim Preserve INAC_array(7, entry_record)	'This resizes the array based on the number of rows in the Excel File'
	'INAC_array (worker_number, 	entry_record) = previous_worker_number	'The client information is added to the array'
	'INAC_array (case_number,  	entry_record) = MAXIS_case_number
	'INAC_array (case_member_name, 	entry_record) = case_name
    'INAC_array (appl_date,	entry_record) = application_date
	'INAC_array (inac_date, 	entry_record) = inactive_date
	'INAC_array (trans_status,   entry_record) = TRUE
	'INAC_array (trans_conf,   entry_record) = ""
	'entry_record = entry_record + 1			'This increments to the next entry in the array'
''	excel_row = excel_row + 1
'Loop


'For each worker_number in INAC_array
	'Establishing values for each case in the array of cases
	'previous_worker_number	= INAC_array(worker_number, item)
	'MAXIS_case_number		= INAC_array(case_number, item)
	'case_name				= INAC_array(case_member_name, item)
	'application_date 		= INAC_array(appl_date, item)
	'inactive_date       	= INAC_array(inac_date, item)
	'transfer_case 			= INAC_array(trans_status, item)
	'transfer_confirmed 		= INAC_array(trans_conf, item)
	'MsgBox previous_worker_number

	'IF transfer_case_action = FALSE THEN
	'	INAC_array(transfer_case, i) = "Excluded Worker"
	'	INAC_array(transfer_confirmed, i) = False
	'END IF
'DO
	IF transfer_case_action = TRUE THEN
	msgbox  transfer_case_action
	'Sets variable for all of the Excel stuff

		EMWriteScreen MAXIS_case_number, 18, 43
    	CALL navigate_to_MAXIS_screen ("SPEC", "XFER")
		'EMWriteScreen maxis_case_number, 18, 43 'MAXIS_case_number'
		TRANSMIT
		Call write_value_and_transmit("x", 7, 16) 'This should have us in SPEC/XWKR'
		EMReadScreen panel_check, 4, 2, 55
		MsgBox panel_check
		EMReadScreen prev_worker, 7, 18, 28
		'MsgBox prev_wor
		PF9
		MsgBox "PF9"
		EMWriteScreen "X127CCL", 18, 61
		CALL clear_line_of_text(18, 74)
		MsgBox "writing"
		TRANSMIT
		'msgbox "where am I"
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
		excel_row = excel_row + 1	'increments the excel row so we don't overwrite our data
	ELSE
		IF transfer_case_action = FLASE THEN
			msgbox  previous_worker_number & " / " & transfer_case_action
			Excel_row = Excel_row + 1
			action_completed = False
	        End if
        'End if
	End if
	'NEXT
	'Export data to Excel
		ObjExcel.Cells(excel_row, 6).Value = trim(transfer_case_action)
		objExcel.cells(excel_row, 7).Value = trim(action_completed)
	   '' Excel_row = Excel_row + 1
LOOP UNTIL previous_worker_number = ""


script_end_procedure("Success! The list is complete. Please review the cases that appear to be in error.")
