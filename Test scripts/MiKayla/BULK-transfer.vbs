'Required for statistical purposes===============================================================================
name_of_script = "BULK - TRANSFER CASE BACK.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 300                               'manual run time, per line, in seconds
STATS_denomination = "I"       'I is for each ITEM
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
CALL changelog_update("06/21/2018", "Updated with requested enhancements.", "MiKayla Handley, Hennepin County")
CALL changelog_update("05/18/2018", "Updated coordinates for writing stats in excel.", "MiKayla Handley, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------Dialog
BeginDialog info_dialog, 0, 0, 266, 115, "BULK - TRANSFER BACK"
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse:", select_a_file_button
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  EditBox 15, 50, 180, 15, file_selection_path
  Text 15, 70, 230, 15, "Select the Excel file that contains your information by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 250, 85, "Using this script:"
EndDialog
'Connects to MAXIS
EMConnect ""
back_to_self
EMWriteScreen "________", 18, 43

'dialog and dialog DO...Loop
Do
    'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    'Show initial dialog
    Do
    	Dialog info_dialog
    	If ButtonPressed = cancel then stopscript
    	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
    Loop until ButtonPressed = OK and file_selection_path <> ""
    If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in


'Call check_for_MAXIS(False)
'back_to_SELF
'Call navigate_to_MAXIS_screen("SPEC", "XFER")

''Opening the Excel file
'Set objExcel = CreateObject("Excel.Application")
'objExcel.Visible = True
'Set objWorkbook = objExcel.Workbooks.Add()
'objExcel.DisplayAlerts = True
'
''Name for the current sheet'
'ObjExcel.ActiveSheet.Name = "Case information"

'Excel headers and formatting the columns
'------------------------------------------------------IEVC'

objExcel.Cells(1, 1).Value     = "CASE NUMBER" 'maxis_case_number
objExcel.Cells(1, 2).Value     = "PREVIOUS XWKR"  'prev_worker
objExcel.Cells(1, 3).Value     = "Case Name" '
objExcel.Cells(1, 4).Value     = "FROM" '

objExcel.Cells(1, 6).Value     = "ACTION COMPLETED"
objExcel.Cells(1, 7).Value     = "PRIV"
FOR i = 1 to 8		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	'ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'This bit freezes the top row of the Excel sheet for better use ability when there is a lot of information
ObjExcel.ActiveSheet.Range("A2").Select
objExcel.ActiveWindow.FreezePanes = True

excel_row = 2

DO
	MAXIS_case_number = ObjExcel.Cells(excel_row, 1).Value
	MAXIS_case_number = trim(MAXIS_case_number)
	IF MAXIS_case_number = "" THEN EXIT DO
	CALL navigate_to_MAXIS_screen ("SPEC", "XFER")
	   EMWriteScreen maxis_case_number, 18, 43 'MAXIS_case_number'
	   TRANSMIT
	   EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets identified, and will not be updated in MMIS
	 	If PRIV_check = "PRIV" then
	   		ObjExcel.Cells(excel_row, 7).Value = "PRIV"
	   Else
   	   	    Call write_value_and_transmit("x", 7, 16)
	   	    'This should have us in SPEC/XWKR'
	   	    EMReadScreen panel_check, 4, 2, 55
	   	    'MsgBox panel_check
	   	    EMReadScreen prev_worker, 7, 18, 28
	   	    'MsgBox prev_worker

	   	    PF9
	   	    'MsgBox "PF9"

	   	    EMWriteScreen "X127EH8", 18, 61
		    CALL clear_line_of_text(18, 74)
	   	    'MsgBox "writing"
	   	    TRANSMIT
	   	    'msgbox "where am I"
	   END IF

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

	If PRIV_check = "PRIV" then action_completed = FALSE
	transfer_case = ObjExcel.Cells(excel_row, 5).Value
	objExcel.Cells(excel_row, 6).Value = action_completed	'Adds worker number to Excel
	excel_row = excel_row + 1	'increments the excel row so we don't overwrite our data
	'blanking out variables
	'maxis_case_number = "" TRANSFERRING AND SERVICING WORKERS MUST BE FROM SAME COUNTY
Loop
'Formatting the column width.
FOR i = 1 to 8
	objExcel.Columns(i).AutoFit()
NEXT

script_end_procedure("Success!")
