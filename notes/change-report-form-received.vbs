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
call changelog_update("01/16/2019", "Updated dialog box to match the form received.", "MiKayla Handley, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


'THE SCRIPT--------------------------------------------------------------------------------------------------------------
'Connect to Bluezone
EMConnect ""
'Grabs Maxis Case number
CALL MAXIS_case_number_finder(MAXIS_case_number)
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 376, 280, "Change Report Form Received"
  EditBox 60, 5, 40, 15, MAXIS_case_number
  EditBox 160, 5, 45, 15, date_received
  EditBox 320, 5, 45, 15, effective_date
  EditBox 50, 35, 315, 15, address_notes
  EditBox 50, 55, 315, 15, household_notes
  EditBox 115, 75, 250, 15, asset_notes
  EditBox 50, 95, 315, 15, vehicles_notes
  EditBox 50, 115, 315, 15, income_notes
  EditBox 50, 135, 315, 15, shelter_notes
  EditBox 50, 155, 315, 15, other_change_notes
  EditBox 60, 180, 305, 15, actions_taken
  EditBox 60, 200, 305, 15, other_notes
  EditBox 70, 220, 295, 15, verifs_requested
  DropListBox 270, 240, 95, 20, "Select One:"+chr(9)+"will continue next month"+chr(9)+"will not continue next month", changes_continue
  CheckBox 10, 245, 140, 10, "Check here to navigate to DAIL/WRIT", tikl_nav_check
  EditBox 75, 260, 85, 15, worker_signature
  ButtonGroup ButtonPressed
	OkButton 260, 260, 50, 15
	CancelButton 315, 260, 50, 15
  Text 5, 10, 50, 10, "Case Number:"
  Text 110, 10, 50, 10, "Effective Date:"
  Text 210, 10, 110, 10, "Date Change Reported/Received:"
  GroupBox 5, 25, 365, 150, "Changes Reported:"
  Text 15, 40, 30, 10, "Address:"
  Text 15, 60, 35, 10, "HH Comp:"
  Text 15, 80, 95, 10, "Assets (savings or property):"
  Text 15, 100, 30, 10, "Vehicles:"
  Text 15, 120, 30, 10, "Income:"
  Text 15, 140, 25, 10, "Shelter:"
  Text 15, 160, 20, 10, "Other:"
  Text 10, 185, 45, 10, "Action Taken:"
  Text 10, 205, 45, 10, "Other Notes:"
  Text 10, 225, 60, 10, "Verifs Requested:"
  Text 10, 265, 60, 10, "Worker Signature:"
  Text 180, 245, 90, 10, "The changes client reports:"
EndDialog
'Shows dialog
DO
	DO
		DO
			DO
    			Dialog Dialog1
				cancel_confirmation
				IF worker_signature = "" THEN MsgBox "You must sign your case note!"
			LOOP UNTIL worker_signature <> ""
			IF IsNumeric(MAXIS_case_number) = FALSE THEN MsgBox "You must type a valid numeric case number."
		LOOP UNTIL IsNumeric(MAXIS_case_number) = TRUE
		IF changes_continue = "Select One:" THEN MsgBox "You Must Select 'The changes client reports field'"
	LOOP UNTIL changes_continue <> "Select One:"
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'Checks Maxis for password prompt
CALL check_for_MAXIS(FALSE)

'THE CASENOTE----------------------------------------------------------------------------------------------------
'Navigates to case note
Call start_a_blank_case_note
CALL write_variable_in_case_note ("--CHANGE REPORTED--")
CALL write_bullet_and_variable_in_case_note("Date Received", date_received)
CALL write_bullet_and_variable_in_case_note("Date Effective", effective_date)
CALL write_bullet_and_variable_in_case_note("Address", address_notes)
CALL write_bullet_and_variable_in_case_note("Household Members", household_notes)
CALL write_bullet_and_variable_in_case_note("Assets", asset_notes)
CALL write_bullet_and_variable_in_case_note("Vehicles", vehicles_notes)
CALL write_bullet_and_variable_in_case_note("Income", income_notes)
CALL write_bullet_and_variable_in_case_note("Shelter", shelter_notes)
CALL write_bullet_and_variable_in_case_note("Other", other_change_notes)
CALL write_bullet_and_variable_in_case_note("Action Taken", actions_taken)
CALL write_bullet_and_variable_in_case_note("Other Notes", other_notes)
CALL write_bullet_and_variable_in_case_note("Verifs Requested", verifs_requested)
IF changes_continue <> "Select One:" THEN CALL write_bullet_and_variable_in_case_note("The changes client reports", changes_continue)
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)

'If we checked to TIKL out, it goes to TIKL and sends a TIKL
IF tikl_nav_check = 1 THEN
	CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
	CALL create_MAXIS_friendly_date(date, 10, 5, 18)
	EMSetCursor 9, 3
END IF
script_end_procedure ("The case note has been created please be sure to send verifications to ECF or case note how the information was received.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------
'--Tab orders reviewed & confirmed----------------------------------------------
'--Mandatory fields all present & Reviewed--------------------------------------
'--All variables in dialog match mandatory fields-------------------------------
'Review dialog names for content and content fit in dialog----------------------
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------
'--CASE:NOTE Header doesn't look funky------------------------------------------
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------
'--MAXIS_background_check reviewed (if applicable)------------------------------
'--PRIV Case handling reviewed -------------------------------------------------
'--Out-of-County handling reviewed----------------------------------------------
'--script_end_procedures (w/ or w/o error messaging)----------------------------
'--BULK - review output of statistics and run time/count (if applicable)--------
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------
'--Incrementors reviewed (if necessary)-----------------------------------------
'--Denomination reviewed -------------------------------------------------------
'--Script name reviewed---------------------------------------------------------
'--BULK - remove 1 incrementor at end of script reviewed------------------------

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------
'--comment Code-----------------------------------------------------------------
'--Update Changelog for release/update------------------------------------------
'--Remove testing message boxes-------------------------------------------------
'--Remove testing code/unnecessary code-----------------------------------------
'--Review/update SharePoint instructions----------------------------------------
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------
'--Complete misc. documentation (if applicable)---------------------------------
'--Update project team/issue contact (if applicable)----------------------------

