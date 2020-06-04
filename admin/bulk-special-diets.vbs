'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK- SPECIAL DIET LIST.vbs"
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("04/21/2020", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone

'Checks to make sure we're in MAXIS
call check_for_MAXIS(True)
'dialog and dialog DO...Loop
Do
	Do
			'The dialog is defined in the loop as it can change as buttons are pressed
			BeginDialog file_select_dialog, 0, 0, 226, 50, "Select the banked months case review file."
  				ButtonGroup ButtonPressed
    			PushButton 175, 10, 40, 15, "Browse...", select_a_file_button
    			OkButton 110, 30, 50, 15
    			CancelButton 165, 30, 50, 15
  				EditBox 5, 10, 165, 15, file_selection_path
			EndDialog
			err_msg = ""
			Dialog file_select_dialog
			cancel_confirmation
			If ButtonPressed = select_a_file_button then
				If file_selection_path <> "" then 'This is handling for if the BROWSE button is pushed more than once'
					objExcel.Quit 'Closing the Excel file that was opened on the first push'
					objExcel = "" 	'Blanks out the previous file path'
				End If
				call file_selection_system_dialog(file_selection_path, ".xlsx") 'allows the user to select the file'
			End If
			If file_selection_path = "" then err_msg = err_msg & vbNewLine & "Use the Browse Button to select the file that has your client data"
			If err_msg <> "" Then MsgBox err_msg
		Loop until err_msg = ""
		If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
		If err_msg <> "" Then MsgBox err_msg
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'resets the case number and footer month/year back to the CM (REVS for current month plus two has is going to be a problem otherwise)
back_to_self
EMwritescreen "________", 18, 43
EMWriteScreen CM_mo_plus_one, 20, 43
EMWriteScreen CM_yr, 20, 46
transmit

case_note_header = "***Recertification Accuracy Update***"
case_note_body = "This client receives a special diet allotment. The Special Diet form was mailed to the client to allow time for a physician to complete the form before the 06/20 recertification is due. If the special diet form is not returned, the MSA will be approved without the special diet allotment. --- CM 23.12 Special Diets need to be verified at recertification even if the special diet form says lifelong or ongoing.--- "
'Required for statistical purposes==========================================================================================
name_of_script = "BULK - INACTIVE TRANSFER BACK.vbs"
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

ObjExcel.Cells(1, 1).Value = "CASE NUMBER"
ObjExcel.Cells(1, 2).Value = "PMI"
ObjExcel.Cells(1, 3).Value = "CASE NAME"
ObjExcel.Cells(1, 4).Value = "PROGRAM"
ObjExcel.Cells(1, 5).Value = "DIET TYPE"
ObjExcel.Cells(1, 6).Value = "DESCRIPTION"
ObjExcel.Cells(1, 7).Value = "WORKER"
ObjExcel.Cells(1, 8).Value = "WORKER NAME"
ObjExcel.Cells(1, 9).Value = "HSS"
ObjExcel.Cells(1, 10).Value = "POPULATION"
ObjExcel.Cells(1, 11).Value = "DATE SENT"
ObjExcel.Cells(1, 12).Value = "HH"
ObjExcel.Cells(1, 13).Value = "NOTES"

FOR i = 1 to 13		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	'ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

CALL check_for_MAXIS(false)
'Now the script adds all the clients on the excel list into an array
excel_row = 2 're-establishing the row to start checking the members for
'entry_record = 0
'This bit freezes the top row of the Excel sheet for better use ability when there is a lot of information
ObjExcel.ActiveSheet.Range("A2").Select
objExcel.ActiveWindow.FreezePanes = True
case_note_action = "TRUE"
DO
    MAXIS_case_number = objExcel.cells(excel_row, 1).Value          're-establishing the case numbers for functions to use
    MAXIS_case_number = trim(MAXIS_case_number)
	PMI_number = objExcel.cells(excel_row, 2).Value          're-establishing the case numbers for functions to use
	PMI_number = trim(PMI_number)
	IF MAXIS_case_number = "" then exit do
	'IF MAXIS_case_number = MAXIS_case_number THEN
	'	IF PMI_number = PMI_number THEN
	'		excel_row = excel_row + 1	'increments the excel row so we don't overwrite our data
	'		case_note_action = "False Duplicate"
	'	ELSE
	'		case_note_action = "Couples"

	'		'CALL navigate_to_MAXIS_screen("STAT", "MEMB")
	'		'EMWriteScreen MAXIS_case_number, 18, 43
	'		'TRANSMIT
	'		'DO
	'		'	EMReadScreen MAXIS_PMI_number, 9, 4, 46
	'		'	MAXIS_PMI_number = trim(MAXIS_PMI_number)
	'		'	IF MAXIS_PMI_number <> PMI_number THEN
	'		' 		TRANSMIT
	'		'	END IF
	'		'LOOP UNTIL MAXIS_PMI_number <> PMI_number
	'		'EMReadScreen MEMB_number, 2, 4, 33
	'	END IF
	'END IF
	'	IF case_note_action = "TRUE" THEN
			CALL navigate_to_MAXIS_screen("CASE", "NOTE")
		    EMWriteScreen MAXIS_case_number, 18, 43
		    TRANSMIT
		    'Checking for privileged
		    EMReadScreen privileged_case, 40, 24, 2
		    IF InStr(privileged_case, "PRIVILEGED") <> 0 THEN
		    	case_note_action = "FALSE PRIV"
		    	'excel_row = excel_row + 1
		    ELSE
		    	PF9
				call write_variable_in_CASE_NOTE("***Recertification Accuracy Update***")
			    call write_variable_in_CASE_NOTE("This client(s) receives a special diet allotment(s). The Special Diet form(s)")
                call write_variable_in_CASE_NOTE("were mailed to allow time for a physician to complete the forms before the")
                call write_variable_in_CASE_NOTE("05/20 recertification is due. If the special diet forms are not returned,")
                call write_variable_in_CASE_NOTE("the MSA will be approved without the special diet allotments.")
                call write_variable_in_CASE_NOTE("---")
                call write_variable_in_CASE_NOTE("CM 23.12 Special Diets need to be verified at recertification even if the")
                call write_variable_in_CASE_NOTE("special diet form says lifelong or ongoing.")
                call write_variable_in_CASE_NOTE("---")
            	call write_variable_in_CASE_NOTE (Worker_signature)
				PF3
		    END IF
		'END IF
		'IF case_note_action = "Couples" THEN
		'		CALL navigate_to_MAXIS_screen("CASE", "NOTE")
		'		EMWriteScreen MAXIS_case_number, 18, 43
		'		TRANSMIT
		'		'Checking for privileged
		'		EMReadScreen privileged_case, 40, 24, 2
		'		IF InStr(privileged_case, "PRIVILEGED") <> 0 THEN
		'			case_note_action = "FALSE PRIV"
		'			'excel_row = excel_row + 1
		'		ELSE
		'			PF9
		'			call write_variable_in_CASE_NOTE("***Recertification Accuracy Update***")
		'			call write_variable_in_CASE_NOTE("These clients receive special diet allotment(s). The Special Diet forms were")
		'			call write_variable_in_CASE_NOTE("mailed to both clients today to allow time for a physician to complete the  ")
		'			call write_variable_in_CASE_NOTE("forms before the 05/20 recertification is due. If the special diet forms are")
		'			call write_variable_in_CASE_NOTE("not returned, the MSA will be approved without the special diet allotments. ")
		'			call write_variable_in_CASE_NOTE("---")
		'			call write_variable_in_CASE_NOTE("CM 23.12 Special Diets need to be verified at recertification even if the   ")
		'			call write_variable_in_CASE_NOTE("special diet form says lifelong or ongoing.                                 ")
		'			call write_variable_in_CASE_NOTE("---")
		'			call write_variable_in_CASE_NOTE (Worker_signature)
		'			PF3
		'		END IF
		'END IF
	'Export data to Excel
	objExcel.cells(excel_row, 13).Value = trim(case_note_action)
	excel_row = excel_row + 1	'increments the excel row so we don't overwrite our data
LOOP UNTIL MAXIS_case_number = ""

STATS_counter = STATS_counter - 1  'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! The list is complete. Please review the cases that appear to be in error.")
