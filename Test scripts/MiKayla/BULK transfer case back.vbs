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
BeginDialog stats_dialog, 0, 0, 266, 85, "Transfer case back"
  ButtonGroup ButtonPressed
    OkButton 150, 65, 50, 15
    CancelButton 205, 65, 50, 15
  Text 15, 25, 135, 10, "This script will gather match information."
  GroupBox 10, 10, 245, 50, "About this script:"
  Text 20, 40, 225, 10, " Please shut down your VGO (not pause it), and press OK to continue."
EndDialog

'Connects to MAXIS
EMConnect ""

Do
	dialog stats_dialog
    cancel_confirmation
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Starting the query start time (for the query runtime at the end)
query_start_time = timer
Call check_for_MAXIS(False)
back_to_SELF
Call navigate_to_MAXIS_screen("REPT", "IEVC")

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Name for the current sheet'
ObjExcel.ActiveSheet.Name = "Case information"

'Excel headers and formatting the columns
'------------------------------------------------------IEVC'
objExcel.Cells(1, 1).Value     = "X1 NUMBER" 'x_number
objExcel.Cells(1, 2).Value     = "CASE NUMBER" 'maxis_case_number
objExcel.Cells(1, 3).Value     = "CLIENT NAME" 'client_name
objExcel.Cells(1, 4).Value     = "APPL DATE" 'appl_date
objExcel.Cells(1, 5).Value     = "INAC DATE" 'inac_date
objExcel.Cells(1, 6).Value     = "TRANSFERED TO"  'spec_xfer_worker

For excel_row = 1 to 6
	objExcel.Cells(excel_row).Font.Bold = True
Next
'This bit freezes the top row of the Excel sheet for better use ability when there is a lot of information
ObjExcel.ActiveSheet.Range("A2").Select
objExcel.ActiveWindow.FreezePanes = True

'Sets variable for all of the Excel stuff
excel_row = 2

DO
    CALL navigate_to_MAXIS_screen ("SPEC", "XFER")
    EMWriteScreen "x", 7, 16
    TRANSMIT
    PF9
    EMReadScreen spec_xfer_worker, 7, 18, 28
    EMWriteScreen spec_xfer_worker 18, 61
    TRANSMIT
    EMReadScreen worker_check, 9, 24, 2
    IF worker_check = "SERVICING" THEN
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
	objExcel.Cells(excel_row, 6).Value = spec_xfer_worker	'Adds worker number to Excel
	excel_row = excel_row + 1	'increments the excel row so we don't overwrite our data


Loop until x_number = ""
'Formatting the column width.
FOR i = 1 to 23
	objExcel.Columns(i).AutoFit()
NEXT

script_end_procedure("Success! The spreadsheet has all requested information.")
