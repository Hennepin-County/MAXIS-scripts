'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - GET CASE STATUS.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 60                      'manual run time in seconds
STATS_denomination = "M"       			   'M is for each CASE
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
call changelog_update("05/29/2019", "Updated script to work with new BOBI query.", "Ilse Ferris, Hennepin County")
call changelog_update("01/31/2019", "Added functionality to change Defer FSET funds field if coded incorrectly on STAT/WREG.", "Ilse Ferris, Hennepin County")
call changelog_update("05/23/2018", "Added code to write in client name if presenting as a PRIV case on initial spreadsheet.", "Ilse Ferris, Hennepin County")
call changelog_update("03/30/2018", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------The script
'CONNECTS TO BlueZone
EMConnect ""

'dialog and dialog DO...Loop
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 301, 100, "ADMIN - ADD POPULATION"
  ButtonGroup ButtonPressed
    PushButton 15, 40, 60, 15, "Browse...", select_a_file_button
  EditBox 80, 40, 205, 15, file_selection_path
  ButtonGroup ButtonPressed
    OkButton 190, 80, 50, 15
    CancelButton 245, 80, 50, 15
  Text 40, 60, 230, 10, "This script should be used when a list of cases needs a population added."
  Text 15, 15, 275, 20, "Select the Excel file that contains your information by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 285, 70, "Using this script:"
EndDialog

'dialog and dialog DO...Loop

Do
    err_msg = ""
    dialog Dialog1
    cancel_without_confirmation
    If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
    If trim(file_selection_path) = "" then err_msg = err_msg & vbcr & "* Select a file to continue."
    If err_msg <> "" Then MsgBox err_msg
Loop until err_msg = ""

'Opening today's list
Call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file

excel_row = 2

Do
	'Grabs the case number
	basket_number = trim(objExcel.cells(excel_row, 1).value)
	If basket_number = "" then exit do

    If basket_number = "X127EF8" then suggested_population = "1800"
    If basket_number = "X127EF9" then suggested_population = "1800"
    If basket_number = "X127EG9" then suggested_population = "1800"
    If basket_number = "X127EG0" then suggested_population = "1800"

    If basket_number = "X127ED8" then suggested_population = "Adults"
    If basket_number = "X127EE1" then suggested_population = "Adults"
    If basket_number = "X127EE2" then suggested_population = "Adults"
    If basket_number = "X127EE3" then suggested_population = "Adults"
    If basket_number = "X127EE4" then suggested_population = "Adults"
    If basket_number = "X127EE5" then suggested_population = "Adults"
    If basket_number = "X127EE6" then suggested_population = "Adults"
    If basket_number = "X127EE7" then suggested_population = "Adults"
    If basket_number = "X127EG4" then suggested_population = "Adults"
    If basket_number = "X127EH8" then suggested_population = "Adults"
    If basket_number = "X127EJ1" then suggested_population = "Adults"
    If basket_number = "X127EL1" then suggested_population = "Adults"
    If basket_number = "X127EL2" then suggested_population = "Adults"
    If basket_number = "X127EL3" then suggested_population = "Adults"
    If basket_number = "X127EL4" then suggested_population = "Adults"
    If basket_number = "X127EL5" then suggested_population = "Adults"
    If basket_number = "X127EL6" then suggested_population = "Adults"
    If basket_number = "X127EL7" then suggested_population = "Adults"
    If basket_number = "X127EL8" then suggested_population = "Adults"
    If basket_number = "X127EL9" then suggested_population = "Adults"
    If basket_number = "X127EN1" then suggested_population = "Adults"
    If basket_number = "X127EN2" then suggested_population = "Adults"
    If basket_number = "X127EN3" then suggested_population = "Adults"
    If basket_number = "X127EN4" then suggested_population = "Adults"
    If basket_number = "X127EN5" then suggested_population = "Adults"
    If basket_number = "X127EN7" then suggested_population = "Adults"
    If basket_number = "X127EP6" then suggested_population = "Adults"
    If basket_number = "X127EP7" then suggested_population = "Adults"
    If basket_number = "X127EP8" then suggested_population = "Adults"
    If basket_number = "X127EQ1" then suggested_population = "Adults"
    If basket_number = "X127EQ3" then suggested_population = "Adults"
    If basket_number = "X127EQ4" then suggested_population = "Adults"
    If basket_number = "X127EQ5" then suggested_population = "Adults"
    If basket_number = "X127EQ8" then suggested_population = "Adults"
    If basket_number = "X127EQ9" then suggested_population = "Adults"
    If basket_number = "X127EX1" then suggested_population = "Adults"
    If basket_number = "X127EX2" then suggested_population = "Adults"
    If basket_number = "X127EX3" then suggested_population = "Adults"
    If basket_number = "X127EX7" then suggested_population = "Adults"
    If basket_number = "X127EX8" then suggested_population = "Adults"
    If basket_number = "X127EX9" then suggested_population = "Adults"
    If basket_number = "X127F3D" then suggested_population = "Adults"
    If basket_number = "X127F3P" then suggested_population = "Adults"   'MA-EPD Adults Basket

    If basket_number = "X127FE7" then suggested_population = "DWP"
    If basket_number = "X127FE8" then suggested_population = "DWP"
    If basket_number = "X127FE9" then suggested_population = "DWP"

    If basket_number = "X127EP8" then suggested_population = "EGA"
    If basket_number = "X127EQ2" then suggested_population = "EGA"

    If basket_number = "X127ES1" then suggested_population = "Families"
    If basket_number = "X127ES2" then suggested_population = "Families"
    If basket_number = "X127ES3" then suggested_population = "Families"
    If basket_number = "X127ES4" then suggested_population = "Families"
    If basket_number = "X127ES5" then suggested_population = "Families"
    If basket_number = "X127ES6" then suggested_population = "Families"
    If basket_number = "X127ES7" then suggested_population = "Families"
    If basket_number = "X127ES8" then suggested_population = "Families"
    If basket_number = "X127ES9" then suggested_population = "Families"
    If basket_number = "X127ET1" then suggested_population = "Families"
    If basket_number = "X127ET2" then suggested_population = "Families"
    If basket_number = "X127ET3" then suggested_population = "Families"
    If basket_number = "X127ET4" then suggested_population = "Families"
    If basket_number = "X127ET5" then suggested_population = "Families"
    If basket_number = "X127ET6" then suggested_population = "Families"
    If basket_number = "X127ET7" then suggested_population = "Families"
    If basket_number = "X127ET8" then suggested_population = "Families"
    If basket_number = "X127ET9" then suggested_population = "Families"
    If basket_number = "X127F4E" then suggested_population = "Families"
    If basket_number = "X127F3H" then suggested_population = "Families"
    If basket_number = "X127FB7" then suggested_population = "Families"
    If basket_number = "X127EZ1" then suggested_population = "Families"
    If basket_number = "X127EZ3" then suggested_population = "Families"
    If basket_number = "X127EZ4" then suggested_population = "Families"
    If basket_number = "X127EZ6" then suggested_population = "Families"
    If basket_number = "X127EZ7" then suggested_population = "Families"
    If basket_number = "X127EZ8" then suggested_population = "Families"
    If basket_number = "X127F3K" then suggested_population = "Families"  'MA-EPD FAD Basket

    If basket_number = "X127EZ2" then suggested_population = "FAD GRH"

    If basket_number = "X127EG5" then suggested_population = "Housing Supports"
    If basket_number = "X127FG3" then suggested_population = "Housing Supports"
    If basket_number = "X127EH2" then suggested_population = "Housing Supports"
    If basket_number = "X127EJ4" then suggested_population = "Housing Supports"
    If basket_number = "X127EJ7" then suggested_population = "Housing Supports"
    If basket_number = "X127EK5" then suggested_population = "Housing Supports"
    If basket_number = "X127EM1" then suggested_population = "Housing Supports"
    If basket_number = "X127EM8" then suggested_population = "Housing Supports"
    If basket_number = "X127EP4" then suggested_population = "Housing Supports"

    If basket_number = "X127EH1" then suggested_population = "LTC+"
    If basket_number = "X127EH3" then suggested_population = "LTC+"
    If basket_number = "X127EH4" then suggested_population = "LTC+"
    If basket_number = "X127EH5" then suggested_population = "LTC+"
    If basket_number = "X127EH6" then suggested_population = "LTC+"
    If basket_number = "X127EH7" then suggested_population = "LTC+"
    If basket_number = "X127EJ8" then suggested_population = "LTC+"
    If basket_number = "X127EK1" then suggested_population = "LTC+"
    If basket_number = "X127EK2" then suggested_population = "LTC+"
    If basket_number = "X127EK3" then suggested_population = "LTC+"
    If basket_number = "X127EK4" then suggested_population = "LTC+"
    If basket_number = "X127EK6" then suggested_population = "LTC+"
    If basket_number = "X127EK7" then suggested_population = "LTC+"
    If basket_number = "X127EK8" then suggested_population = "LTC+"
    If basket_number = "X127EK9" then suggested_population = "LTC+"
    If basket_number = "X127EM9" then suggested_population = "LTC+"
    If basket_number = "X127EN6" then suggested_population = "LTC+"
    If basket_number = "X127EP5" then suggested_population = "LTC+"
    If basket_number = "X127EP9" then suggested_population = "LTC+"
    If basket_number = "X127EZ5" then suggested_population = "LTC+"
    If basket_number = "X127F3F" then suggested_population = "LTC+"
    If basket_number = "X127FE5" then suggested_population = "LTC+"
    If basket_number = "X127FH4" then suggested_population = "LTC+"
    If basket_number = "X127FH5" then suggested_population = "LTC+"
    If basket_number = "X127FI2" then suggested_population = "LTC+"
    If basket_number = "X127FI7" then suggested_population = "LTC+"
    'Contacted Case Mgt
    If basket_number = "X127FG6" then suggested_population = "LTC+"           '"Kristen Kasem"
    If basket_number = "X127FG7" then suggested_population = "LTC+"           '"Kristen Kasem"
    If basket_number = "X127EM3" then suggested_population = "LTC+"           '"True L. or Gina G."
    If basket_number = "X127EM4" then suggested_population = "LTC+"            '"True L. or Gina G."
    If basket_number = "X127EW7" then suggested_population = "LTC+"            '"Kimberly Hill"
    If basket_number = "X127EW8" then suggested_population = "LTC+"            '"Kimberly Hill"
    If basket_number = "X127FF4" then suggested_population = "LTC+"            '"Alyssa Taylor"
    If basket_number = "X127FF5" then suggested_population = "LTC+"            '"Alyssa Taylor"

    If basket_number = "X127EH9" then suggested_population = "LTH"
    If basket_number = "X127EJ1" then suggested_population = "LTH"
    If basket_number = "X127EM2" then suggested_population = "LTH"
    If basket_number = "X127FE6" then suggested_population = "LTH"

    If basket_number = "X127FA5" then suggested_population = "YET"
    If basket_number = "X127FA6" then suggested_population = "YET"
    If basket_number = "X127FA7" then suggested_population = "YET"
    If basket_number = "X127FA8" then suggested_population = "YET"
    If basket_number = "X127FB1" then suggested_population = "YET"
    If basket_number = "X127FA9" then suggested_population = "YET"

	objExcel.Cells(excel_row, 2).value = suggested_population
    excel_row = excel_row + 1
    STATS_counter = stats_counter + 1
LOOP
STATS_counter = STATS_counter - 1 'removes one from the count since 1 is counted at the beginning

script_end_procedure("All done.")
