'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - INCARCERATION.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 90           'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THIS SCRIPT IS BEING USED IN A WORKFLOW SO DIALOGS ARE NOT NAMED
'DIALOGS MAY NOT BE DEFINED AT THE BEGINNING OF THE SCRIPT BUT WITHIN THE SCRIPT FILE

'THE SCRIPT CODE-----------------------------------------------------------------------------------------------------

'Connects to BLUEZONE
EMConnect ""

'Grabs the MAXIS case number
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Shows and defines the FIRST dialog box
BeginDialog , 0, 0, 166, 85, "Incarceration"
  EditBox 80, 5, 75, 15, MAXIS_case_number
  EditBox 80, 25, 75, 15, hh_member
  EditBox 80, 45, 25, 15, month_benefit
  EditBox 115, 45, 25, 15, year_benefit
  ButtonGroup ButtonPressed
    OkButton 65, 65, 45, 15
    CancelButton 115, 65, 45, 15
  Text 5, 10, 70, 10, "Maxis Case Number:"
  Text 5, 50, 70, 10, "Benefit Month/Year:"
  Text 5, 30, 70, 10, "HH Member Number:"
EndDialog
DO
	Dialog 					'Calling a dialog without a assigned variable will call the most recently defined dialog
	cancel_confirmation
	IF isnumeric(MAXIS_case_number)= FALSE THEN MsgBox "You must enter a valid case number!"
LOOP UNTIL Isnumeric(MAXIS_case_number) = TRUE

CALL navigate_to_MAXIS_screen("stat", "faci")
	EMReadScreen panel_max_check, 1, 2, 78
	IF panel_max_check = "5" THEN
		script_end_procedure ("This case has reached the maximum amount of FACI panels.  Please review your case, delete an appropriate FACI panel, and run the script again.")
	'ELSE
		'EMWriteScreen "nn", 20, 79
		'transmit
	END IF

'Shows and defines the MAIN dialog
BeginDialog , 0, 0, 451, 200, "Incarceration"
  EditBox 85, 10, 85, 15, MAXIS_case_number
  EditBox 280, 10, 75, 15, hh_member
  EditBox 85, 40, 85, 15, start_date
  EditBox 280, 40, 110, 15, incarceration_location
  EditBox 95, 70, 90, 15, date_out
  DropListBox 255, 70, 130, 15, "Select One..."+chr(9)+"County Correctional Facility"+chr(9)+"Non-County Adult Correctional", faci_type
  ComboBox 150, 100, 100, 15, "Click here to enter info"+chr(9)+"Client"+chr(9)+"AREP"+chr(9)+"Jail Roster Search"+chr(9)+"Child Support Officer"+chr(9)+"Probation Officer"+chr(9)+"Social Worker", info_recd
  EditBox 345, 100, 90, 15, po_info
  EditBox 55, 125, 135, 15, actions_taken
  EditBox 285, 125, 150, 15, other_notes
  CheckBox 90, 150, 80, 15, "Update STAT/FACI", update_faci_checkbox
  CheckBox 205, 150, 235, 15, "Navigate to DAIL/WRIT to create a TIKL to check for release date", tikl_checkbox
  EditBox 90, 175, 95, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 310, 175, 50, 15
    CancelButton 365, 175, 50, 15
  Text 15, 180, 75, 10, "Sign Your Case Note:"
  Text 95, 85, 70, 15, "(Ex: MM/DD/YYYY)"
  Text 5, 130, 50, 10, "Actions Taken:"
  Text 5, 45, 80, 10, "Incarceration Start Date:"
  Text 205, 45, 75, 10, "Incarceration Location:"
  Text 15, 15, 70, 10, "Maxis Case Number:"
  Text 5, 105, 140, 10, "Info Rec'd Via (or type in info rec'd source):"
  Text 195, 15, 85, 10, "HH Member Incarcerated:"
  Text 270, 105, 75, 10, "Probation Officer Info:"
  Text 85, 55, 80, 15, "(Ex: MM/DD/YYYY)"
  Text 205, 130, 80, 10, "Other Notes/Comments:"
  Text 280, 25, 40, 15, "(Ex: 01, 02)"
  Text 205, 75, 45, 10, "Facility Type:"
  Text 5, 75, 85, 20, "Anticipated Release Date: (Leave Blank if Unknown)"
EndDialog
DO
	err_msg = ""
	Dialog  					'Calling a dialog without a assigned variable will call the most recently defined dialog
		IF ButtonPressed = 0 THEN StopScript
		IF info_recd = "Click here to enter info" THEN err_msg = err_msg & vbCr & "You must select how the incarceration info was received!"
		IF faci_type = "Select One..." THEN err_msg = err_msg & vbCr & "You must select a facility type!"
		IF IsNumeric(MAXIS_case_number) = FALSE THEN err_msg = err_msg & vbCr & "You must type a valid numeric case number."
		IF start_date = "" OR (start_date <> "" AND IsDate(start_date) = False) THEN err_msg = err_msg & vbCr & "You must enter a date in a MM/DD/YYYY format!"
		IF actions_taken = "" THEN err_msg = err_msg & vbCr & "You must enter actions taken!"
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "You must sign your case note!"
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""

'Checks Maxis for password prompt
CALL check_for_MAXIS(False)


'IF the update STAT/FACI checkbox was checked, then the script will navigate to that panel and updated it
IF update_faci_checkbox = checked THEN
	CALL navigate_to_MAXIS_screen("stat", "faci")
	EMReadScreen panel_max_check, 1, 2, 78
	IF panel_max_check = "5" THEN
		script_end_procedure ("This case has reached the maximum amount of FACI panels.  Please review your case, delete an appropriate FACI panel, and run the script again.")
	ELSE
		EMWriteScreen "nn", 20, 79
		transmit
	END IF

	'Writes the facility name in the Facility Name field
	EMWriteScreen incarceration_location, 6, 43

	'Writes the 68 or 69 in the Facility Type field
	IF faci_type = "County Correctional Facility" THEN EMWriteScreen "68", 7, 43
	IF faci_type = "Non-County Adult Correctional" THEN EMWriteScreen "69", 7, 43

	'Writes the N in the FS Eligible Y/N field
	EMWriteScreen "N", 8, 43

	'Writes the Incarceration Start Date in the Date In field
	CALL create_MAXIS_friendly_date_with_YYYY(start_date, 0, 14, 47)

	'Writes the Anticipted Release Date in the Date Out field if there is a date out
	IF date_out <> "" THEN CALL create_MAXIS_friendly_date_with_YYYY(date_out, 0, 14, 71)
END IF

'Opens a new case note
start_a_blank_case_note

'Writes the case note
CALL write_variable_in_case_note ("===Incarceration Reported===")
CALL write_bullet_and_variable_in_case_note("HH Member Incarcerated", hh_member)
CALL write_bullet_and_variable_in_case_note("Incarceration Start Date", start_date)
CALL write_bullet_and_variable_in_case_note("Incarceration Location/Facility", incarceration_location)
CALL write_bullet_and_variable_in_case_note("Anticipated Release Date", date_out)
CALL write_bullet_and_variable_in_case_note("Facility Type", faci_type)
CALL write_bullet_and_variable_in_case_note("Info Rec'd Via", info_recd)
CALL write_bullet_and_variable_in_case_note("Probation Officer Info", po_info)
CALL write_bullet_and_variable_in_case_note("Actions Taken", actions_taken)
CALL write_bullet_and_variable_in_case_note("Other Comments/Notes", other_notes)
IF update_faci_checkbox = checked THEN CALL write_variable_in_case_note("* Updated STAT/FACI")
IF tikl_checkbox = checked THEN CALL write_variable_in_case_note("* TIKL'd to check for release date")
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)

'If worker checked to TIKL out, it goes to DAIL WRIT
IF tikl_checkbox = checked THEN
	CALL navigate_to_MAXIS_screen("DAIL","WRIT")
	CALL create_MAXIS_friendly_date(date, 10, 5, 18)
	EMSetCursor 9, 3
	EMSendKey "Check status of HH member " & hh_member & "'s incarceration at " & incarceration_location & ". Incarceration Start Date was " & start_date & "."
END IF

script_end_procedure("")
