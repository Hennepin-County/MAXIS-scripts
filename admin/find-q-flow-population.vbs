'Required for statistical purposes==========================================================================================
name_of_script = "ADMIN - FIND Q FLOW POPULATION.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 90          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
    IF on_the_desert_island = TRUE Then
        FuncLib_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\desert-island\MASTER FUNCTIONS LIBRARY.vbs"
        Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
        Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
        text_from_the_other_script = fso_command.ReadAll
        fso_command.Close
        Execute text_from_the_other_script
    ELSEIF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
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
call changelog_update("04/22/2022", "Suggested Q-Flow populations updated to support LTC+ and Housing Supports roll out.", "Ilse Ferris, Hennepin County")
call changelog_update("03/03/2022", "Updated baskets including new EGA basket, X127EP3.", "Ilse Ferris, Hennepin County")
call changelog_update("12/16/2020", "Added multi-case search functionalty.", "Ilse Ferris, Hennepin County")
call changelog_update("12/09/2020", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT--------------------------------------------------------------------------------------------------
'CONNECTING TO MAXIS & GRABBING THE CASE NUMBER
EMConnect ""
Call Check_for_MAXIS(False)
'CALL MAXIS_case_number_finder(MAXIS_case_number)
end_msg = "Case Numbers reviewed: "

Do
    MAXIS_case_number = ""
    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog dialog1, 0, 0, 116, 60, "Case # Dialog"
      EditBox 65, 5, 45, 15, MAXIS_case_number
      ButtonGroup ButtonPressed
        OkButton 25, 25, 40, 15
        CancelButton 70, 25, 40, 15
      Text 10, 10, 50, 10, "Case Number:"
      CheckBox 20, 45, 90, 10, "Checking multiple cases.", multi_case_checkbox
    EndDialog

    Do
    	DO
    		err_msg = ""
    	    dialog dialog1
          	cancel_without_confirmation
          	IF IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

    Call navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", is_this_priv)
    If is_this_priv = TRUE then script_end_procedure("Privileged case, cannot access/update. The script will now end.")

    '----------------------------------------------------------------------------------------------------Adding suggested Q-Flow Ticketing population for follow up work. needed during the COVID-19 PEACETIME STATE OF EMERGENCY
    EmReadscreen basket_number, 7, 21, 14    'Reading basket number on CASE/CURR
    suggested_population = ""                'Blanking this out. Will default to no suggestions if x number is not in this this.

    '----------------------------------------------------------------------------------------------------ADS
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
    If basket_number = "X127EJ4" then suggested_population = "LTC+"
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

    If trim(suggested_population) = "" then suggested_population = "No suggestions available"
    stats_counter = stats_counter + 1
    msgbox "Q Flow Population: " & suggested_population
    end_msg = end_msg & MAXIS_case_number & ", "
    If multi_case_checkbox = 0 then exit do
Loop

end_msg = trim(end_msg)  'trims excess spaces of end_msg
If right(end_msg, 1) = "," THEN end_msg = left(end_msg, len(end_msg) - 1)

stats_counter = stats_counter - 1 'removing extra count
script_end_procedure(end_msg)

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------04/26/2022
'--Tab orders reviewed & confirmed----------------------------------------------04/26/2022
'--Mandatory fields all present & Reviewed--------------------------------------04/26/2022
'--All variables in dialog match mandatory fields-------------------------------04/26/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------04/26/2022---------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------04/26/2022---------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------04/26/2022---------------------N/A
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------04/26/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------04/26/2022---------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------04/26/2022
'--Out-of-County handling reviewed----------------------------------------------04/26/2022---------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------04/26/2022---------------------N/A
'--BULK - review output of statistics and run time/count (if applicable)--------04/26/2022---------------------N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------04/26/2022
'--Incrementors reviewed (if necessary)-----------------------------------------04/26/2022
'--Denomination reviewed -------------------------------------------------------04/26/2022
'--Script name reviewed---------------------------------------------------------04/26/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------04/26/2022

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete-----------------------------------------04/26/2022
'--comment Code-----------------------------------------------------------------04/26/2022
'--Update Changelog for release/update------------------------------------------04/26/2022
'--Remove testing message boxes-------------------------------------------------04/26/2022
'--Remove testing code/unnecessary code-----------------------------------------04/26/2022
'--Review/update SharePoint instructions----------------------------------------04/26/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------04/26/2022---------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------04/26/2022
'--Complete misc. documentation (if applicable)---------------------------------04/26/2022
'--Update project team/issue contact (if applicable)----------------------------04/26/2022
