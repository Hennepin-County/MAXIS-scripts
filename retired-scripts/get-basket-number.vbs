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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("04/26/2021", "Updated ADAD and FAD basket numbers to reflect current baskets.", "Ilse Ferris, Hennepin County")
call changelog_update("03/27/2020", "Updated basket numbers to reflect current baskets.", "Ilse Ferris, Hennepin County")
call changelog_update("09/12/2019", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display


'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone

'dialog and dialog DO...Loop
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 221, 50, "Select the case list source file"
    ButtonGroup ButtonPressed
    PushButton 175, 10, 40, 15, "Browse...", select_a_file_button
    OkButton 110, 30, 50, 15
    CancelButton 165, 30, 50, 15
    EditBox 5, 10, 165, 15, file_selection_path
EndDialog
Do
    'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    'Show initial dialog
    Do
        err_msg = ""
    	Dialog Dialog1
    	If ButtonPressed = cancel then stopscript
    	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
        If file_selection_path = "" then err_msg = err_msg & vbNewLine & "Use the Browse Button to select the file that has your client data"
        If err_msg <> "" Then MsgBox err_msg
    Loop until ButtonPressed = OK and file_selection_path <> ""
    If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 126, 50, "Select the excel row to start"
  EditBox 75, 5, 40, 15, excel_row_to_start
  ButtonGroup ButtonPressed
    OkButton 10, 25, 50, 15
    CancelButton 65, 25, 50, 15
  Text 10, 10, 60, 10, "Excel row to start:"
EndDialog

do
    dialog Dialog1
    If buttonpressed = 0 then stopscript								'loops until all errors are resolved
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

excel_row = excel_row_to_start

back_to_self
EMWriteScreen CM_mo, 20, 43
EMWriteScreen CM_yr, 20, 46

pop_col = 1
basket_col = 2
case_num_col = 3 

DO
    'Grabs the case number
	MAXIS_case_number = objExcel.cells(excel_row, case_num_col).value
    If MAXIS_case_number = "" then exit do
	
    Call navigate_to_MAXIS_screen("CASE", "CURR")
    EMReadScreen CURR_panel_check, 4, 2, 55
	'If CURR_panel_check <> "CURR" then ObjExcel.Cells(excel_row, basket_col).Value = ""

    EMReadScreen basket, 7, 21, 14
	ObjExcel.Cells(excel_row, basket_col).Value = basket

    '----------------------------------------------------------------------------------------------------ADS
    If basket = "X127EF8" then population_type = "1800"     '1800
    If basket = "X127EF9" then population_type = "1800"     '1800
    If basket = "X127EG9" then population_type = "1800"     '1800
    If basket = "X127EG0" then population_type = "1800"     '1800
    
    If basket = "X127EJ6" then population_type = "ADS"	
    If basket = "X127FE5" then population_type = "ADS"	
    If basket = "X127EK1" then population_type = "ADS"	
    If basket = "X127EK2" then population_type = "ADS"	
    If basket = "X127EJ7" then population_type = "ADS"	
    If basket = "X127EJ8" then population_type = "ADS"	
    If basket = "X127EK3" then population_type = "ADS"	
    If basket = "X127EH6" then population_type = "ADS"	
    If basket = "X127EM1" then population_type = "ADS"	
    If basket = "X127FI7" then population_type = "ADS"	
    If basket = "X127EK4" then population_type = "ADS"	
    If basket = "X127EK5" then population_type = "ADS"	
    If basket = "X127FH5" then population_type = "ADS"	
    If basket = "X127EK6" then population_type = "ADS"	
    If basket = "X127EK9" then population_type = "ADS"	
    If basket = "X127EM7" then population_type = "ADS"	
    If basket = "X127FI2" then population_type = "ADS"	
    If basket = "X127FG3" then population_type = "ADS"	
    If basket = "X127EM8" then population_type = "ADS"	
    If basket = "X127EM9" then population_type = "ADS"	
    If basket = "X127EJ4" then population_type = "ADS"	
    If basket = "X127EJ5" then population_type = "ADS"	    
    If basket = "X127EH1" then population_type = "ADS"
    If basket = "X127EH7" then population_type = "ADS"
    If basket = "X127EH2" then population_type = "ADS"
    If basket = "X127EH3" then population_type = "ADS"
    If basket = "X127EN6" then population_type = "ADS"
    If basket = "X127FH4" then population_type = "ADS"
    If basket = "X127EP3" then population_type = "ADS"
    If basket = "X127EP4" then population_type = "ADS"
    If basket = "X127EP5" then population_type = "ADS"
    If basket = "X127EP9" then population_type = "ADS"
    If basket = "X127F3U" then population_type = "ADS"
    If basket = "X127F3V" then population_type = "ADS"
    
    '----------------------------------------------------------------------------------------------------Adults     
    If basket = "X127ED8" then population_type = "Adults"
    If basket = "X127EG4" then population_type = "Adults"
    If basket = "X127EH8" then population_type = "Adults"
    If basket = "X127EN5" then population_type = "Adults"
    If basket = "X127EP6" then population_type = "Adults"
    If basket = "X127EP7" then population_type = "Adults"
    If basket = "X127EQ3" then population_type = "Adults"
    If basket = "X127EE1" then population_type = "Adults"
    If basket = "X127EE2" then population_type = "Adults"
    If basket = "X127EE3" then population_type = "Adults"
    If basket = "X127EE4" then population_type = "Adults"
    If basket = "X127EE5" then population_type = "Adults"
    If basket = "X127EE6" then population_type = "Adults"
    If basket = "X127EE7" then population_type = "Adults"
    If basket = "X127EL1" then population_type = "Adults"
    If basket = "X127EL2" then population_type = "Adults"
    If basket = "X127EL3" then population_type = "Adults"
    If basket = "X127EL4" then population_type = "Adults"
    If basket = "X127EL5" then population_type = "Adults"
    If basket = "X127EL6" then population_type = "Adults"
    If basket = "X127EL7" then population_type = "Adults"
    If basket = "X127EL8" then population_type = "Adults"
    If basket = "X127EL9" then population_type = "Adults"
    If basket = "X127EN1" then population_type = "Adults"
    If basket = "X127EN2" then population_type = "Adults"
    If basket = "X127EN3" then population_type = "Adults"
    If basket = "X127EN4" then population_type = "Adults"
    If basket = "X127EN7" then population_type = "Adults"
    If basket = "X127EQ1" then population_type = "Adults"
    If basket = "X127EQ4" then population_type = "Adults"
    If basket = "X127EQ5" then population_type = "Adults"
    If basket = "X127EQ8" then population_type = "Adults"
    If basket = "X127EQ9" then population_type = "Adults"
    
    '----------------------------------------------------------------------------------------------------Families
    If basket ="X127ET9" then population_type = "Families"
    If basket ="X127ET8" then population_type = "Families"
    If basket ="X127ES8" then population_type = "Families"
    If basket ="X127ES3" then population_type = "Families"
    If basket ="X127ES1" then population_type = "Families"
    If basket ="X127ES2" then population_type = "Families"
    If basket ="X127ES4" then population_type = "Families"
    If basket ="X127ES5" then population_type = "Families"
    If basket ="X127ES6" then population_type = "Families"
    If basket ="X127ES7" then population_type = "Families"
    If basket ="X127ES9" then population_type = "Families"
    If basket ="X127ET1" then population_type = "Families"
    If basket ="X127ET2" then population_type = "Families"
    If basket ="X127ET3" then population_type = "Families"
    If basket ="X127ET4" then population_type = "Families"
    If basket ="X127ET5" then population_type = "Families"
    If basket ="X127ET6" then population_type = "Families"
    If basket ="X127ET7" then population_type = "Families"
    
    '----------------------------------------------------------------------------------------------------DWP
    If basket = "X127FE7" then population_type = "DWP"
    If basket = "X127FE8" then population_type = "DWP"
    If basket = "X127FE9" then population_type = "DWP"
    
    '----------------------------------------------------------------------------------------------------YET
    If basket = "X127FA5" then population_type = "YET"
    If basket = "X127FA6" then population_type = "YET"
    If basket = "X127FA7" then population_type = "YET"
    If basket = "X127FA8" then population_type = "YET"
    If basket = "X127FB1" then population_type = "YET"
    If basket = "X127F3S" then population_type = "YET"
    If basket = "X127FA9" then population_type = "YET"
        
    '----------------------------------------------------------------------------------------------------METS
    If basket = "X127F4A" then population_type = "METS"
    If basket = "X127F4B" then population_type = "METS"
    If basket = "X127FI1" then population_type = "METS"
    If basket = "X127FI3" then population_type = "METS"
    If basket = "X127EX4" then population_type = "METS"
    If basket = "X127EX5" then population_type = "METS"
    If basket = "X127FF1" then population_type = "METS"
    If basket = "X127FF2" then population_type = "METS"
    If basket = "X127FH3" then population_type = "METS"
    If basket = "X127FI6" then population_type = "METS"
    If basket = "X127EN8" then population_type = "METS"
    If basket = "X127EN9" then population_type = "METS"
    If basket = "X127EQ6" then population_type = "METS"
    If basket = "X127EQ7" then population_type = "METS"
    If basket = "X127EP1" then population_type = "METS"
    If basket = "X127EP2" then population_type = "METS"
    If basket = "X127FE2" then population_type = "METS"
    If basket = "X127FE3" then population_type = "METS"
    If basket = "X127FG5" then population_type = "METS"
    If basket = "X127FG9" then population_type = "METS"
    If basket = "X127F3E" then population_type = "METS"
    If basket = "X127F3J" then population_type = "METS"
    If basket = "X127F3N" then population_type = "METS"
    
    '-----------------------------------------------------------------------------------------------------SHELTER/EA
    If basket = "X127LE1" then population_type = "Shelter/EA"
    If basket = "X127SH1" then population_type = "Shelter/EA"
    If basket = "X127AN1" then population_type = "Shelter/EA"
    If basket = "X127EHD" then population_type = "Shelter/EA"
    If basket = "X127EA0" then population_type = "Shelter/EA"
    If basket = "X127EAK" then population_type = "Shelter/EA"

    '----------------------------------------------------------------------------------------------------Speciality HC
    If basket = "X127FF6" then population_type = "Speciality HC"    'HCMC, MHC, North Memorial
    If basket = "X127FF7" then population_type = "Speciality HC"
    If basket = "X127ER7" then population_type = "Speciality HC" 
    If basket = "X127FF8" then population_type = "Speciality HC"
    If basket = "X127FF9" then population_type = "Speciality HC"
    
    If basket = "X127F3F" then population_type = "MA-EPD"       'MA-EPD
    If basket = "X127F3K" then population_type = "MA-EPD"       'MA-EPD
    If basket = "X127F3P" then population_type = "MA-EPD"       'MA-EPD
    
    If basket = "X127FG6" then population_type = "LTC"      'Contracted Case Management      
    If basket = "X127FG7" then population_type = "LTC"      'Contracted Case Management
    If basket = "X127EM3" then population_type = "LTC"      'Contracted Case Management
    If basket = "X127EM4" then population_type = "LTC"      'Contracted Case Management
    If basket = "X127EW7" then population_type = "LTC"      'Contracted Case Management
    If basket = "X127EW8" then population_type = "LTC"      'Contracted Case Management
    If basket = "X127NP0" then population_type = "LTC"      'Contracted Case Management
    If basket = "X127NPC" then population_type = "LTC"      'Contracted Case Management
    If basket = "X127FF4" then population_type = "LTC"      'Contracted Case Management
    If basket = "X127FF5" then population_type = "LTC"      'Contracted Case Management

    If basket = "X127FG1" then population_type = "IV-E"     'IV-E
    If basket = "X127EW6" then population_type = "IV-E"     'IV-E
    If basket = "X1274EC" then population_type = "IV-E"     'IV-E
    If basket = "X127FG2" then population_type = "IV-E"     'IV-E
    If basket = "X127EW4" then population_type = "IV-E"     'IV-E
    
    ObjExcel.Cells(excel_row, pop_col).Value = population_type

    MAXIS_case_number = ""
    basket = ""
    population_type = ""
    excel_row = excel_row + 1
    STATS_counter = STATS_counter + 1

LOOP UNTIL objExcel.Cells(excel_row, case_num_col).value = ""	'looping until the list of cases to check for recert is complete

STATS_counter = STATS_counter - 1 'removes one from the count since 1 is counted at the beginning (because counting :p)
script_end_procedure("Success! The Excel file now has been updated. Please review the blank case statuses that remain.")