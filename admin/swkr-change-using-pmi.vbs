'Required for statistical purposes===============================================================================
name_of_script = "BULK - SWKR CHANGE USING PMI.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 180                               'manual run time in seconds
STATS_denomination = "C"       'C is for each Case
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

'creating an array of cases to update
excel_row = 2
Do 
	'establishing variables for case to be updated
	client_PMI = objExcel.cells(excel_row, 2).value
	client_PMI = trim(client_PMI)
	first_name = objExcel.cells(excel_row, 3).value
	first_name = trim(first_name)
	last_name = objExcel.cells(excel_row, 4).value
	last_name = trim(last_name)
	phone_ext = objExcel.cells(excel_row, 5).value
	phone_ext = trim(phone_ext)
	
	IF client_PMI = "" then exit do
	'trims all the 0's off of the PMI number 
	Do 
		if left(client_PMI, 1) = "0" then client_PMI = right(client_PMI, len(client_PMI) -1)
	Loop until left(client_PMI, 1) <> "0"
	
	back_to_self
	EMWriteScreen "________", 18, 43					'clears case number
	call navigate_to_MAXIS_screen("pers", "____")
	EMWriteScreen client_PMI, 15, 36
	Transmit
	EMReadscreen PMI_confirmation, 10, 8, 71
	PMI_confirmation = trim(PMI_confirmation)
	'msgbox PMI_confirmation
	If PMI_confirmation <> client_PMI then
		msgbox client_PMI & " does not match client. Process manually."
	Else 	
		EMWriteScreen "x", 8, 5		
		Transmit
		
		'chekcing for an active case
		MAXIS_row = 10
		Do 
			EMReadscreen open_case, 5, MAXIS_row, 53
			open_case = trim(open_case)
			If open_case = "" then
				EMReadscreen MAXIS_case_number, 8, MAXIS_row, 6
				MAXIS_case_number = trim(MAXIS_case_number) 
		 		EMWriteScreen "x", MAXIS_row, 4
				Transmit
				Exit do
			Else 
				MAXIS_row = MAXIS_row + 1
			END IF 
		LOOP until MAXIS_row = 19
		If MAXIS_row = 19 then msgbox "Unable to find an open case for " & client_PMI & vbnewline & "excel row: " & excel_row
		
		'navigating to the SWKR panel, and ensuring that we are in that panel
		back_to_self
		EMWriteScreen MAXIS_case_number, 18, 43
		Call navigate_to_MAXIS_screen("STAT", "SWKR")
		'msgbox "what's happening?"
		'Checking for privileged
		EMReadScreen privileged_case, 40, 24, 2
		IF InStr(privileged_case, "PRIVILEGED") <> 0 THEN 
			msgbox "PRIV case " & client_PMI & vbnewline & "Excel row: " & excel_row
			privileged_array = privileged_array & client_PMI & "~~~"
			EMWriteScreen "________", 18, 43
			Transmit
		ELSE
			Do 
				EMReadscreen SWKR_panel_check, 4, 2, 47
				If SWKR_panel_check <> "SWKR" then call write_value_and_transmit("SWKR", 20, 71)
			Loop until SWKR_panel_check = "SWKR"
				
			EMReadScreen needs_SWKR_panel, 1, 2, 73
			If needs_SWKR_panel = "0" then 
				call write_value_and_transmit("NN", 20, 79)
			ELSE 
				PF9
			END IF 
				
			EMReadScreen edit_check, 2, 24, 2
			If edit_check <> "  " then 
				msgbox client_PMI & " at excel row: " & excel_row & " cannot be updated. Log the PMI number and update manually."
			ELSE 
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
				PF3
			END IF
		END if 
	END IF
	back_to_self
	excel_row = excel_row + 1
	client_PMI = ""
	MAXIS_case_number = ""
	STATS_counter = STATS_counter + 1    'adds one instance to the stats counter
Loop until excel_row = 280

IF privileged_array <> "" THEN 
	privileged_array = replace(privileged_array, "~~~", vbCr)
	MsgBox "The script could not generate a CASE NOTE for the following cases..." & vbCr & privileged_array
END IF

STATS_counter = STATS_counter - 1  'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
msgbox STATS_counter
script_end_procedure("Success! The addresses have been updated!")
