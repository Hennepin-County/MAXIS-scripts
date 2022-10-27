
'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - GET CASE STATUS.vbs"
start_time = timer
STATS_counter = 0                     'sets the stats counter at one
STATS_manualtime = 25                 'manual run time in seconds
STATS_denomination = "C"       	      'C is for each CASE

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

'Custom function for this script----------------------------------------------------------------------------------------------------
FUNCTION get_case_status
	back_to_self
	EMWriteScreen MAXIS_case_number, 18, 43

	Call navigate_to_MAXIS_screen("CASE", "CURR")
	EMReadScreen CURR_panel_check, 4, 2, 55
	If CURR_panel_check <> "CURR" then ObjExcel.Cells(excel_row, 5).Value = ""

	EMReadScreen case_status, 8, 8, 9
	case_status = trim(case_status)
	ObjExcel.Cells(excel_row, 6).Value = case_status
	IF case_status = " - Syste" THEN case_status = "PRIV" 'doesnt access the screen'
	MAXIS_case_number = ""
	excel_row = excel_row + 1
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
'End of function----------------------------------------------------------------------------------------------------

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone

'dialog and dialog DO...Loop
Do
	Do
			'The dialog is defined in the loop as it can change as buttons are pressed
			'-------------------------------------------------------------------------------------------------DIALOG
			BeginDialog Dialog1, 0, 0, 181, 45, "Select the file to check case status"
  			ButtonGroup ButtonPressed
    		PushButton 0, 5, 40, 15, "Browse:", select_a_file_button
  			EditBox 45, 5, 130, 15, file_selection_path
  			ButtonGroup ButtonPressed
    		OkButton 105, 25, 35, 15
    		CancelButton 140, 25, 35, 15
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
			If isnumeric(excel_row_to_start) = False then msgbox "Enter a valid numeric row to start."
			If file_selection_path = "" then err_msg = err_msg & vbNewLine & "Use the Browse Button to select the file that has your client data"
			If err_msg <> "" Then MsgBox err_msg
		Loop until err_msg = ""
		If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
		If err_msg <> "" Then MsgBox err_msg
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Excel headers and formatting the columns
objExcel.Cells(1, 1).Value = "X1 NUMBER"
objExcel.Cells(1, 2).Value = "CASE NBR"
objExcel.Cells(1, 3).Value = "CLIENT NAME"
objExcel.Cells(1, 4).Value = "DAIL TYPE"
objExcel.Cells(1, 5).Value = "DAIL MONTH"
objExcel.Cells(1, 6).Value = "DAIL MESSAGE"
objExcel.Cells(1, 7).Value = "STATUS"

FOR i = 1 to 6		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Sets variable for all of the Excel stuff
excel_row = 2
'resets the case number and footer month/year back to the CM (REVS for current month plus two has is going to be a problem otherwise)
back_to_self
EMWriteScreen CM_mo, 20, 43
EMWriteScreen CM_yr, 20, 46
transmit

STATS_counter = 1
'Zeroing out variables
active_status = 0
pending_status = 0
inactive_status = 0
rein_status = 0

	objExcel.Columns(3).EntireColumn.Insert
'FOR column = 2 to 2
	objExcel.Columns(3).EntireColumn.Insert
'NEXT
'Columns(B).EntireColumn.Insert
'Sub sbInsertingColumns()
'Inserting a Column at Column B
'sht.Range("B1").EntireColumn.Insert
'objExcel.Columns().EntireColumn.Insert
'ObjExcel.Cells("b1").EntireColumn.Insert
'Range("B1").EntireColumn.Insert
'4832'



excel_row = 2
Do
	'Grabs the case number
	MAXIS_case_number = objExcel.cells(excel_row, 2).value
	If MAXIS_case_number = "" then exit do
	get_case_status
LOOP UNTIL objExcel.Cells(excel_row, 2).value = ""	'looping until the list of cases to check for recert is complete

STATS_counter = STATS_counter - 1 'removes one from the count since 1 is counted at the beginning


script_end_procedure("Success! The Excel file now has been updated with case status.")
