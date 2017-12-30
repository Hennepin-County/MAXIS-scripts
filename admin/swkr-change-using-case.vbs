'Required for statistical purposes===============================================================================
name_of_script = "BULK - SWKR CHANGE USING CASE.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 150         'manual run time in seconds
STATS_denomination = "C"       'C is for each Case
'END OF stats block==============================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("12/12/2017", "Updated for Touchstone project.", "Ilse Ferris, Hennepin County"
call changelog_update("07/31/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone

'dialog and dialog DO...Loop	
Do
	Do
		'The dialog is defined in the loop as it can change as buttons are pressed 
		BeginDialog file_select_dialog, 0, 0, 221, 50, "Select the file that contains the SWRK information."
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

excel_row = 2
back_to_self
Do 
	'establishing variables for case to be updated
	MAXIS_case_number = objExcel.cells(excel_row, 2).value
	If trim(MAXIS_case_number) = "" then exit do

	call navigate_to_MAXIS_screen("STAT", "SWKR")
	PF9
	'Clears SWKR address
	'EMWriteScreen "___________________________________", 6, 32	
	EMWriteScreen "______________________", 8, 32
	EMWriteScreen "______________________", 9, 32
	EMWriteScreen "_______________", 10, 32
	EMWriteScreen "__", 10, 54
	EMWriteScreen "_______", 10, 63
	''Clears SWKR phone number
	'EMWriteScreen "___", 12, 34
	'EMWriteScreen "___", 12, 40
	'EMWriteScreen "___", 12, 44
	'EMWriteScreen "____", 12, 54
	
	'Writes in the address for Touchstone 
	'EMWriteScreen first_name & " " & last_name & "/People Inc.", 6, 32	
	EMWriteScreen "2312 Snelling Ave", 8, 32
	EMWriteScreen "Minneapolis", 10, 32
	EMWriteScreen "MN", 10, 54
	EMWriteScreen "55404", 10, 63
	
	'Clears SWKR phone number
	'EMWriteScreen "612", 12, 34
	'EMWriteScreen "230", 12, 40
	'EMWriteScreen "6270", 12, 44
	'EMWriteScreen phone_ext, 12, 54
	'EMWriteScreen "Y", 15, 63			'coding notices to be sent to SWKR
	'msgbox "confirm case: " & client_PMI & " at excel row " & excel_row
	Transmit
	Transmit
	transmit
		
	'Adds the new address inforamtion to the Excel sheet
	EMReadScreen SWKR_name, 34, 6, 32
	EMReadScreen addr_one, 22, 8, 32
	EMReadScreen addr_two, 22, 9, 32
	EMReadScreen city, 15, 10, 32
	EMReadScreen state, 2, 10, 54
	EMReadScreen zip_code, 7, 10, 63
	EMReadScreen phone_number, 26, 12, 32
	If phone_number = "( ___ ) ___ ____ Ext: ____" then phone_number = ""
	EMReadScreen send_notice, 1, 15, 63
	
	ObjExcel.Cells(excel_row, 4).Value = replace(swkr_name, "_", "")
	ObjExcel.Cells(excel_row, 5).Value = replace(addr_one, "_", "")
	ObjExcel.Cells(excel_row, 6).Value = replace(addr_two, "_", "")
	ObjExcel.Cells(excel_row, 7).Value = replace(city, "_", "")
	ObjExcel.Cells(excel_row, 8).Value = replace(state, "_", "")
	ObjExcel.Cells(excel_row, 9).Value = replace(zip_code, "_", "")
	ObjExcel.Cells(excel_row, 10).Value = replace(phone_number, "_", "")
	ObjExcel.Cells(excel_row, 11).Value = replace(send_notice, "_", "")
	
	PF3	'To leave SWKR panel, background will trigger
	
	start_a_blank_CASE_NOTE
	Call write_value_and_transmit("SWKR panel updated w/ Touchstone's new address")
	Call write_value_and_transmit("Notices were going to the old address on University Avenue. New address is: 2312 Snelling Avenue, Minneapolis, MN. 55404")
	Call write_value_and_transmit("---")
	Call write_value_and_transmit("I. Ferris/QI team")
	
	excel_row = excel_row + 1
	MAXIS_case_number = ""
	STATS_counter = STATS_counter + 1    'adds one instance to the stats counter
Loop

STATS_counter = STATS_counter - 1  'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
msgbox STATS_counter
script_end_procedure("Success! The addresses have been updated!")
