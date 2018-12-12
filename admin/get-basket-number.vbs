'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - GET BASKET NUMBER.vbs" 'BULK script that creates a list of cases that require an interview, and the contact phone numbers'
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

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone

'The dialog is defined in the loop as it can change as buttons are pressed 
BeginDialog file_select_dialog, 0, 0, 221, 50, "Select the case list source file"
    ButtonGroup ButtonPressed
    PushButton 175, 10, 40, 15, "Browse...", select_a_file_button
    OkButton 110, 30, 50, 15
    CancelButton 165, 30, 50, 15
    EditBox 5, 10, 165, 15, file_selection_path
EndDialog

BeginDialog excel_row_dialog, 0, 0, 126, 50, "Select the excel row to start"
  EditBox 75, 5, 40, 15, excel_row_to_start
  ButtonGroup ButtonPressed
    OkButton 10, 25, 50, 15
    CancelButton 65, 25, 50, 15
  Text 10, 10, 60, 10, "Excel row to start:"
EndDialog

'dialog and dialog DO...Loop	
Do
    'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    'Show initial dialog
    Do
        err_msg = ""
    	Dialog file_select_dialog
    	If ButtonPressed = cancel then stopscript
    	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
        If file_selection_path = "" then err_msg = err_msg & vbNewLine & "Use the Browse Button to select the file that has your client data"
        If err_msg <> "" Then MsgBox err_msg
    Loop until ButtonPressed = OK and file_selection_path <> ""
    If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

do 
    dialog excel_row_dialog
    If buttonpressed = 0 then stopscript								'loops until all errors are resolved
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

excel_row = excel_row_to_start 

back_to_self
EMwritescreen "________", 18, 43
EMWriteScreen CM_mo, 20, 43
EMWriteScreen CM_yr, 20, 46

case_num_col = 3
basket_col = 1
pop_col = 2

DO  
    'Grabs the case number
	MAXIS_case_number = objExcel.cells(excel_row, case_num_col).value
    If MAXIS_case_number = "" then exit do
	back_to_self
	EMWriteScreen "________", 18, 43
	EMWriteScreen MAXIS_case_number, 18, 43
	
    Call navigate_to_MAXIS_screen("CASE", "CURR")
    EMReadScreen CURR_panel_check, 4, 2, 55
	If CURR_panel_check <> "CURR" then ObjExcel.Cells(excel_row, basket_col).Value = ""
    
    EMReadScreen basket, 7, 21, 14
	ObjExcel.Cells(excel_row, basket_col).Value = basket
        
    'ADS----------------------------------------------------------------------------------------------------
    If basket = "X127EJ6" then population_type = "ADS"		'1: Central/NE
    If basket = "X127FE5" then population_type = "ADS"		'1: Central/NE
    If basket = "X127EK3" then population_type = "ADS"		'1: Central/NE
    If basket = "X127EK1" then population_type = "ADS"		'1: Central/NE
    If basket = "X127EK2" then population_type = "ADS"		'1: Central/NE
    If basket = "X127EJ7" then population_type = "ADS"		'1: Central/NE
    If basket = "X127EJ8" then population_type = "ADS"		'1: Central/NE
    If basket = "X127EJ5" then population_type = "ADS"		'1: Central/NE
    
    If basket = "X127EH6" then population_type = "ADS"		'2: North Mpls	
    If basket = "X127EM1" then population_type = "ADS"		'2: North Mpls
    If basket = "X127FE1" then population_type = "ADS"		'2: North Mpls
    If basket = "X127FI7" then population_type = "ADS"		'2: North Mpls
    If basket = "X127FH3" then population_type = "ADS"		'2: North Mpls
    If basket = "X127F3E" then population_type = "ADS"		'2: North Mpls
    If basket = "X127F3J" then population_type = "ADS"		'2: North Mpls
    If basket = "X127F3N" then population_type = "ADS"		'2: North Mpls
    If basket = "X127FI6" then population_type = "ADS"		'2: North Mpls
    
    If basket = "X127EK9" then population_type = "ADS"		'3: Northwest
    If basket = "X127FH5" then population_type = "ADS"		'3: Northwest
    If basket = "X127EK5" then population_type = "ADS"		'3: Northwest
    If basket = "X127EN7" then population_type = "ADS"		'3: Northwest
    If basket = "X127EK6" then population_type = "ADS"		'3: Northwest
    If basket = "X127EK4" then population_type = "ADS"		'3: Northwest
    If basket = "X127EN6" then population_type = "ADS"		'3: Northwest
    If basket = "X127EL1" then population_type = "ADS"		'3: Northwest
    If basket = "X127ER6" then population_type = "ADS"		'3: Northwest    
    If basket = "X127EP8" then population_type = "ADS"		'3: Northwest
    If basket = "X127EQ3" then population_type = "ADS"		'3: Northwest
    If basket = "X127FG9" then population_type = "ADS"		'3: Northwest
    If basket = "X127FI3" then population_type = "ADS"		'3: Northwest
    
    If basket = "X127EM7" then population_type = "ADS"		'4: South Mpls
    If basket = "X127FI2" then population_type = "ADS"		'4: South Mpls
    If basket = "X127FG3" then population_type = "ADS"		'4: South Mpls
    If basket = "X127EM8" then population_type = "ADS"		'4: South Mpls
    If basket = "X127EM9" then population_type = "ADS"		'4: South Mpls
    If basket = "X127EJ4" then population_type = "ADS"		'4: South Mpls
    
    If basket = "X127EH1" then population_type = "ADS"		'5: South Suburban
    If basket = "X127EH7" then population_type = "ADS"		'5: South Suburban
    If basket = "X127EH2" then population_type = "ADS"		'5: South Suburban
    If basket = "X127EH3" then population_type = "ADS"		'5: South Suburban
    If basket = "X127FH4" then population_type = "ADS"		'5: South Suburban
    If basket = "X127FI1" then population_type = "ADS"		'5: South Suburban
    
    If basket = "X127EP3" then population_type = "ADS"		'6: West
    If basket = "X127EP4" then population_type = "ADS"		'6: West
    If basket = "X127EP5" then population_type = "ADS"		'6: West
    If basket = "X127EP9" then population_type = "ADS"		'6: West
    
    'adults----------------------------------------------------------------------------------------------------         
    If basket = "X127EJ9" then population_type = "Adults"		'1: Central/NE
    If basket = "X127EQ8" then population_type = "Adults"		'1: Central/NE
    If basket = "X127EQ9" then population_type = "Adults"		'1: Central/NE
    If basket = "X127EE2" then population_type = "Adults"		'1: Central/NE
    If basket = "X127EE3" then population_type = "Adults"		'1: Central/NE
    If basket = "X127EE4" then population_type = "Adults"		'1: Central/NE
    If basket = "X127EE5" then population_type = "Adults"		'1: Central/NE
    If basket = "X127EE6" then population_type = "Adults"		'1: Central/NE
    If basket = "X127EE7" then population_type = "Adults"		'1: Central/NE
    If basket = "X127ER1" then population_type = "Adults"		'1: Central/NE
    If basket = "X127ER2" then population_type = "Adults"		'1: Central/NE
    If basket = "X127ER3" then population_type = "Adults"		'1: Central/NE
    If basket = "X127ER4" then population_type = "Adults"		'1: Central/NE
    If basket = "X127EG8" then population_type = "Adults"		'1: Central/NE
    If basket = "X127ER5" then population_type = "Adults"		'1: Central/NE
    If basket = "X127EG5" then population_type = "Adults"		'1: Central/NE
    If basket = "X127FH2" then population_type = "Adults"		'1: Central/NE
    If basket = "X127EHD" then population_type = "Adults"		'1: Central/NE
        
    If basket = "X127EL8" then population_type = "Adults"		'2: North Mpls
    If basket = "X127EL9" then population_type = "Adults"		'2: North Mpls
    If basket = "X127EL2" then population_type = "Adults"		'2: North Mpls
    If basket = "X127EL3" then population_type = "Adults"		'2: North Mpls
    If basket = "X127EL4" then population_type = "Adults"		'2: North Mpls
    If basket = "X127EL5" then population_type = "Adults"		'2: North Mpls
    If basket = "X127EL6" then population_type = "Adults"		'2: North Mpls
    If basket = "X127EL7" then population_type = "Adults"		'2: North Mpls
    If basket = "X127FG5" then population_type = "Adults"		'2: North Mpls
    
    If basket = "X127EQ1" then population_type = "Adults"		'3: Northwest
    If basket = "X127EF7" then population_type = "Adults"		'3: Northwest
    If basket = "X127EN5" then population_type = "Adults"		'3: Northwest
    If basket = "X127EQ2" then population_type = "Adults"		'3: Northwest
    If basket = "X127EF5" then population_type = "Adults"		'3: Northwest
    If basket = "X127EK7" then population_type = "Adults"		'3: Northwest
    If basket = "X127EF6" then population_type = "Adults"		'3: Northwest
    If basket = "X127EQ5" then population_type = "Adults"		'3: Northwest
    If basket = "X127EK8" then population_type = "Adults"		'3: Northwest
    If basket = "X127EQ4" then population_type = "Adults"		'3: Northwest
    If basket = "X127FH9" then population_type = "Adults"		'3: Northwest
    If basket = "X127EG6" then population_type = "Adults"		'3: Northwest
    
    If basket = "X127ED8" then population_type = "Adults"		'4: South Mpls
    If basket = "X127EH8" then population_type = "Adults"		'4: South Mpls
    If basket = "X127EAJ" then population_type = "Adults"		'4: South Mpls
    If basket = "X127EN1" then population_type = "Adults"		'4: South Mpls
    If basket = "X127EN2" then population_type = "Adults"		'4: South Mpls
    If basket = "X127EN3" then population_type = "Adults"		'4: South Mpls
    If basket = "X127EN4" then population_type = "Adults"		'4: South Mpls
    If basket = "X127ED6" then population_type = "Adults"		'4: South Mpls
    If basket = "X127ED7" then population_type = "Adults"		'4: South Mpls
    If basket = "X127EJ2" then population_type = "Adults"		'4: South Mpls
    If basket = "X127EJ3" then population_type = "Adults"		'4: South Mpls
    If basket = "X127FH1" then population_type = "Adults"		'4: South Mpls
    If basket = "X127FG4" then population_type = "Adults"		'4: South Mpls
    If basket = "X127F3C" then population_type = "Adults"		'4: South Mpls
    If basket = "X127F3G" then population_type = "Adults"		'4: South Mpls
    If basket = "X127F3L" then population_type = "Adults"		'4: South Mpls
    If basket = "X127EJ1" then population_type = "Adults"		'4: South Mpls
    If basket = "X127EH9" then population_type = "Adults"		'4: South Mpls
    If basket = "X127EM2" then population_type = "Adults"		'4: South Mpls
    If basket = "X127FE6" then population_type = "Adults"		'4: South Mpls
    '1800 has been assigned to South Mpls
    If basket = "X127EF8" then population_type = "1800"		'4: South Mpls
    If basket = "X127EF9" then population_type = "1800"		'4: South Mpls
    If basket = "X127EAK" then population_type = "1800"		'4: South Mpls
    If basket = "X127EG9" then population_type = "1800"		'4: South Mpls
    If basket = "X127EG0" then population_type = "1800"		'4: South Mpls
    
    If basket = "X127EE1" then population_type = "Adults"		'5: South Suburban
    If basket = "X127FB2" then population_type = "Adults"		'5: South Suburban
    If basket = "X127EG7" then population_type = "Adults"		'5: South Suburban
    If basket = "X127ED9" then population_type = "Adults"		'5: South Suburban
    If basket = "X127EE0" then population_type = "Adults"		'5: South Suburban
    If basket = "X127EH4" then population_type = "Adults"		'5: South Suburban
    If basket = "X127EH5" then population_type = "Adults"		'5: South Suburban
    If basket = "X127F3D" then population_type = "Adults"		'5: South Suburban
    If basket = "X127FH8" then population_type = "Adults"		'5: South Suburban
    
    If basket = "X127EP6" then population_type = "Adults"		'6: West
    If basket = "X127EP7" then population_type = "Adults"		'6: West
    If basket = "X127EG4" then population_type = "Adults"		'6: West
    If basket = "X127FG8" then population_type = "Adults"		'6: West
    'DWP----------------------------------------------------------------------------------------------------
    If basket = "X127EY8" then population_type = "DWP"		'2: North Mpls		
    If basket = "X127EY9" then population_type = "DWP"		'2: North Mpls
    If basket = "X127EZ1" then population_type = "DWP"		'2: North Mpls
    
    If basket = "X127FE7" then population_type = "DWP"		'4: South Mpls		
    If basket = "X127FE8" then population_type = "DWP"		'4: South Mpls
    If basket = "X127FE9" then population_type = "DWP"		'4: South Mpls
    'Families----------------------------------------------------------------------------------------------------
    If basket = "X127FD4" then population_type = "Families"		'1: Central/NE
    If basket = "X127FD5" then population_type = "Families"		'1: Central/NE
    If basket = "X127EZ5" then population_type = "Families"		'1: Central/NE
    If basket = "X127FD8" then population_type = "Families"		'1: Central/NE
    If basket = "X127EZ8" then population_type = "Families"		'1: Central/NE
    If basket = "X127FH6" then population_type = "Families"		'1: Central/NE
    If basket = "X127FD6" then population_type = "Families"		'1: Central/NE
    If basket = "X127EZ6" then population_type = "Families"		'1: Central/NE
    If basket = "X127EZ7" then population_type = "Families"		'1: Central/NE
    If basket = "X127FD9" then population_type = "Families"		'1: Central/NE
    If basket = "X127FD7" then population_type = "Families"		'1: Central/NE
    If basket = "X127EZ0" then population_type = "Families"		'1: Central/NE
    If basket = "X127EDD" then population_type = "Families"		'1: Central/NE
    
    If basket = "X127ES4" then population_type = "Families"		'2: North Mpls
    If basket = "X127ET2" then population_type = "Families"		'2: North Mpls
    If basket = "X127ET3" then population_type = "Families"		'2: North Mpls
    If basket = "X127FJ2" then population_type = "Families"		'2: North Mpls
    If basket = "X127EX3" then population_type = "Families"		'2: North Mpls
    If basket = "X127ES8" then population_type = "Families"		'2: North Mpls
    If basket = "X127ET1" then population_type = "Families"		'2: North Mpls
    If basket = "X127ES7" then population_type = "Families"		'2: North Mpls
    If basket = "X127EM5" then population_type = "Families"		'2: North Mpls
    If basket = "X127EM6" then population_type = "Families"		'2: North Mpls
    If basket = "X127EZ2" then population_type = "Families"		'2: North Mpls
    If basket = "X127EZ9" then population_type = "Families"		'2: North Mpls
    If basket = "X127ES5" then population_type = "Families"		'2: North Mpls
    If basket = "X127EX2" then population_type = "Families"		'2: North Mpls
    If basket = "X127ES6" then population_type = "Families"		'2: North Mpls
    If basket = "X127EZ4" then population_type = "Families"		'2: North Mpls
    If basket = "X127EZ3" then population_type = "Families"		'2: North Mpls
    If basket = "X127ES9" then population_type = "Families"		'2: North Mpls
    If basket = "X127EX1" then population_type = "Families"		'2: North Mpls
    If basket = "X127FF3" then population_type = "Families"		'2: North Mpls
    If basket = "X127EW7" then population_type = "ADS"		    '2: North Mpls ADS for FAD
    If basket = "X127EW8" then population_type = "Families"		'2: North Mpls
    If basket = "X127EW9" then population_type = "Families"		'2: North Mpls
    
    If basket = "X127EU5" then population_type = "Families"		'3: Northwest
    If basket = "X127EX7" then population_type = "Families"		'3: Northwest
    If basket = "X127F3Y" then population_type = "Families"		'3: Northwest
    If basket = "X127FA3" then population_type = "Families"		'3: Northwest
    If basket = "X127EU6" then population_type = "Families"		'3: Northwest
    If basket = "X127F3S" then population_type = "Families"		'3: Northwest
    If basket = "X127FJ5" then population_type = "Families"		'3: Northwest
    If basket = "X127EY1" then population_type = "Families"		'3: Northwest
    If basket = "X127EY2" then population_type = "Families"		'3: Northwest
    If basket = "X127F3W" then population_type = "Families"		'3: Northwest
    If basket = "X127FA1" then population_type = "Families"		'3: Northwest
    If basket = "X127EU8" then population_type = "Families"		'3: Northwest
    If basket = "X127F3Q" then population_type = "Families"		'3: Northwest
    If basket = "X127EX9" then population_type = "Families"		'3: Northwest
    If basket = "X127FA4" then population_type = "Families"		'3: Northwest
    If basket = "X127BV1" then population_type = "Families"		'3: Northwest
    If basket = "X127F3T" then population_type = "Families"		'3: Northwest
    If basket = "X127FJ1" then population_type = "Families"		'3: Northwest
    If basket = "X127EU9" then population_type = "Families"		'3: Northwest
    If basket = "X127F3X" then population_type = "Families"		'3: Northwest
    If basket = "X127FA2" then population_type = "Families"		'3: Northwest
    If basket = "X127EU7" then population_type = "Families"		'3: Northwest
    If basket = "X127F3R" then population_type = "Families"		'3: Northwest
    If basket = "X127EX8" then population_type = "Families"		'3: Northwest
    If basket = "X127F3Z" then population_type = "Families"		'3: Northwest
    If basket = "X127FJ3" then population_type = "Families"		'3: Northwest
    If basket = "X127FJ4" then population_type = "Families"		'3: Northwest
    If basket = "X127F3V" then population_type = "Families"		'3: Northwest
    If basket = "X127F3U" then population_type = "Families"		'3: Northwest
    
    If basket = "X127EV1" then population_type = "Families"		'4: South Mpls
    If basket = "X127FB9" then population_type = "Families"		'4: South Mpls
    If basket = "X127FC1" then population_type = "Families"		'4: South Mpls
    If basket = "X127EV5" then population_type = "Families"		'4: South Mpls
    If basket = "X127FC2" then population_type = "Families"		'4: South Mpls
    If basket = "X127EV2" then population_type = "Families"		'4: South Mpls
    If basket = "X127EV4" then population_type = "Families"		'4: South Mpls
    If basket = "X127EV3" then population_type = "Families"		'4: South Mpls
    If basket = "X127FB8" then population_type = "Families"		'4: South Mpls
    If basket = "X127FB7" then population_type = "Families"		'4: South Mpls
    
    If basket = "X127ER8" then population_type = "Families"		'5: South Suburban
    If basket = "X127ET4" then population_type = "Families"		'5: South Suburban
    If basket = "X127F3B" then population_type = "Families"		'5: South Suburban
    If basket = "X127ET6" then population_type = "Families"		'5: South Suburban
    If basket = "X127ES1" then population_type = "Families"		'5: South Suburban
    If basket = "X127ES3" then population_type = "Families"		'5: South Suburban
    If basket = "X127FB6" then population_type = "Families"		'5: South Suburban
    If basket = "X127ET8" then population_type = "Families"		'5: South Suburban
    If basket = "X127F3H" then population_type = "Families"		'5: South Suburban
    If basket = "X127F4E" then population_type = "Families"		'5: South Suburban
    If basket = "X127FB4" then population_type = "Families"		'5: South Suburban
    If basket = "X127F3A" then population_type = "Families"		'5: South Suburban
    If basket = "X127F4C" then population_type = "Families"		'5: South Suburban
    If basket = "X127F4F" then population_type = "Families"		'5: South Suburban
    If basket = "X127FB5" then population_type = "Families"		'5: South Suburban
    If basket = "X127F4D" then population_type = "Families"		'5: South Suburban
    If basket = "X127F3M" then population_type = "Families"		'5: South Suburban
    If basket = "X127ET7" then population_type = "Families"		'5: South Suburban
    If basket = "X127FB3" then population_type = "Families"		'5: South Suburban
    If basket = "X127ER9" then population_type = "Families"		'5: South Suburban
    If basket = "X127ET5" then population_type = "Families"		'5: South Suburban
    If basket = "X127ES2" then population_type = "Families"		'5: South Suburban
    If basket = "X127BV3" then population_type = "Families"		'5: South Suburban
    
    If basket = "X127ET9" then population_type = "Families"		'6: West
    If basket = "X127EU4" then population_type = "Families"		'6: West
    If basket = "X127EW2" then population_type = "Families"		'6: West
    If basket = "X127EW3" then population_type = "Families"		'6: West
    If basket = "X127FH7" then population_type = "Families"		'6: West
    If basket = "X127EU1" then population_type = "Families"		'6: West
    If basket = "X127EU3" then population_type = "Families"		'6: West
    If basket = "X127BV2" then population_type = "Families"		'6: West
    If basket = "X127EU2" then population_type = "Families"		'6: West
    'YET----------------------------------------------------------------------------------------------------
    If basket = "X127CCR" then population_type = "YET"		'0: YET only 	
    If basket = "X127CCA" then population_type = "YET"		'0: YET only 
    If basket = "X127FA5" then population_type = "YET"		'0: YET only 
    If basket = "X127FA6" then population_type = "YET"		'0: YET only 
    If basket = "X127FA7" then population_type = "YET"		'0: YET only 
    If basket = "X127FA8" then population_type = "YET"		'0: YET only 
    If basket = "X127FB1" then population_type = "YET"		'0: YET only 
    If basket = "X127FA9" then population_type = "YET"		'0: YET only 
    'METS----------------------------------------------------------------------------------------------------
    If basket = "X127EN8" then population_type = "METS"		'1: Central/NE
    If basket = "X127EN9" then population_type = "METS"		'1: Central/NE
    
    If basket = "X127F4A" then population_type = "METS"		'2: North Mpls
    If basket = "X127F4B" then population_type = "METS"		'2: North Mpls
    
    If basket = "X127EX4" then population_type = "METS"		'3: Northwest
    If basket = "X127EX5" then population_type = "METS"		'3: Northwest
    If basket = "X127FF1" then population_type = "METS"		'3: Northwest
    If basket = "X127FF2" then population_type = "METS"		'3: Northwest
    
    If basket = "X127EQ6" then population_type = "METS"		'4: South Mpls
    If basket = "X127EQ7" then population_type = "METS"		'4: South Mpls
    
    If basket = "X127EP1" then population_type = "METS"		'5: South Suburban
    If basket = "X127EP2" then population_type = "METS"		'5: South Suburban
    
    If basket = "X127FE2" then population_type = "METS"		'6: West
    If basket = "X127FE3" then population_type = "METS"		'6: West
    
    If basket = "X127EM3" then population_type = "LTC"      'LTC Facilities
    If basket = "X127FF4" then population_type = "LTC" 
    If basket = "X127FG6" then population_type = "LTC" 
    If basket = "X127EM4" then population_type = "LTC"  
    If basket = "X127FG7" then population_type = "LTC"
    
    If basket = "X127F3F" then population_type = "MA-EPD"
    If basket = "X127F3K" then population_type = "MA-EPD"
    If basket = "X127F3P" then population_type = "MA-EPD"
    
    If basket = "X127EA0" then population_type = "Families" 'EA'
    
    If basket = "X127FG1" then population_type = "IV-E"
    If basket = "X127EW6" then population_type = "IV-E"
    If basket = "X1274EC" then population_type = "IV-E"
    If basket = "X127FG2" then population_type = "IV-E"
    If basket = "X127EW4" then population_type = "IV-E"
    If basket = "X127EW5" then population_type = "IV-E"
        
    ObjExcel.Cells(excel_row, pop_col).Value = population_type
    
    MAXIS_case_number = ""
    basket = ""
    population_type = ""
    excel_row = excel_row + 1
    STATS_counter = STATS_counter + 1    
    
LOOP UNTIL objExcel.Cells(excel_row, case_num_col).value = ""	'looping until the list of cases to check for recert is complete

STATS_counter = STATS_counter - 1 'removes one from the count since 1 is counted at the beginning (because counting :p)
script_end_procedure("Success! The Excel file now has been updated. Please review the blank case statuses that remain.")