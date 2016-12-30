'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - CHANGE REPORT FORM RECEIVED.vbs"
start_timer = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 90                      'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
'END OF stats block=========================================================================================================

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
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE DIALOG--------------------------------------------------------------------------------------------------------------
BeginDialog crf_received_dialog, 0, 0, 411, 320, "Change Report Form Received"
  EditBox 55, 5, 55, 15, MAXIS_case_number
  EditBox 270, 5, 60, 15, date_received
  EditBox 50, 35, 340, 15, address_notes
  EditBox 75, 55, 315, 15, household_notes
  EditBox 50, 75, 340, 15, savings_notes
  EditBox 50, 95, 340, 15, property_notes
  EditBox 50, 115, 340, 15, vehicles_notes
  EditBox 50, 135, 340, 15, income_notes
  EditBox 45, 155, 345, 15, shelter_notes
  EditBox 40, 175, 350, 15, other
  EditBox 55, 205, 325, 15, actions_taken
  EditBox 50, 225, 330, 15, other_notes
  EditBox 70, 245, 310, 15, verifs_requested
  CheckBox 10, 270, 140, 15, "Check here to navigate to DAIL/WRIT", tikl_nav_check
  DropListBox 275, 270, 95, 20, "Select One..."+chr(9)+"will continue next month"+chr(9)+"will not continue next month", changes_continue
  EditBox 80, 295, 85, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 290, 300, 50, 15
    CancelButton 345, 300, 50, 15
  Text 20, 180, 25, 10, "Other:"
  Text 20, 40, 30, 15, "Address:"
  Text 10, 210, 45, 10, "Action Taken:"
  Text 20, 100, 35, 15, "Property:"
  Text 10, 230, 40, 10, "Other notes:"
  Text 160, 10, 110, 10, "Change Report Form Rec'd Date:"
  Text 10, 250, 60, 10, "Verifs Requested:"
  Text 20, 120, 35, 15, "Vehicles:"
  Text 20, 60, 60, 15, "Household Mbrs:"
  Text 185, 270, 90, 15, "The changes client reports:"
  Text 20, 140, 30, 15, "Income:"
  Text 10, 300, 70, 10, "Sign your case note:"
  Text 5, 10, 50, 10, "Case Number:"
  Text 20, 160, 30, 15, "Shelter:"
  Text 20, 80, 30, 15, "Savings:"
  GroupBox 5, 25, 395, 175, "Changes Reported:"
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------------------
'Connect to Bluezone
EMConnect ""
'Grabs Maxis Case number
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Shows dialog
DO
	DO
		DO
			DO
				Dialog crf_received_dialog
				cancel_confirmation
				IF worker_signature = "" THEN MsgBox "You must sign your case note!"
			LOOP UNTIL worker_signature <> ""
			IF IsNumeric(MAXIS_case_number) = FALSE THEN MsgBox "You must type a valid numeric case number."
		LOOP UNTIL IsNumeric(MAXIS_case_number) = TRUE
		IF changes_continue = "Select One..." THEN MsgBox "You Must Select 'The changes client reports field'"
	LOOP UNTIL changes_continue <> "Select One..."
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'Checks Maxis for password prompt
CALL check_for_MAXIS(FALSE)

'THE CASE NOTE----------------------------------------------------------------------------------------------------
'Navigates to case note
Call start_a_blank_CASE_NOTE
CALL write_variable_in_case_note ("***Change Report Form Received***")
CALL write_bullet_and_variable_in_case_note("Date Form Received", date_received)
CALL write_bullet_and_variable_in_case_note("Address", address_notes)
CALL write_bullet_and_variable_in_case_note("Household Members", household_notes)
CALL write_bullet_and_variable_in_case_note("Savings", savings_notes)
CALL write_bullet_and_variable_in_case_note("Property", property_notes)
CALL write_bullet_and_variable_in_case_note("Vehicles", vehicles_notes)
CALL write_bullet_and_variable_in_case_note("Income", income_notes)
CALL write_bullet_and_variable_in_case_note("Shelter", shelter_notes)
CALL write_bullet_and_variable_in_case_note("Other", other)
CALL write_bullet_and_variable_in_case_note("Action Taken", actions_taken)
CALL write_bullet_and_variable_in_case_note("Other Notes", other_notes)
CALL write_bullet_and_variable_in_case_note("Verifs Requested", verifs_requested)
IF changes_continue <> "select one..." THEN CALL write_bullet_and_variable_in_case_note("The changes client reports", changes_continue)
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)

'If we checked to TIKL out, it goes to TIKL and sends a TIKL
IF tikl_nav_check = 1 THEN
	CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
	CALL create_MAXIS_friendly_date(date, 10, 5, 18)
	EMSetCursor 9, 3
END IF

script_end_procedure("")
