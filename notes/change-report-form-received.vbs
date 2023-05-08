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
call changelog_update("05/02/2023", "Updated field entry validation for dialog box.", "Mark Riegel, Hennepin County")
call changelog_update("01/16/2019", "Updated dialog box to match the form received.", "MiKayla Handley, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


'THE SCRIPT--------------------------------------------------------------------------------------------------------------
'Connect to Bluezone
EMConnect ""
' Checks Maxis for password prompt
CALL check_for_MAXIS(FALSE)
'Grabs Maxis Case number
CALL MAXIS_case_number_finder(MAXIS_case_number)
get_county_code()

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 376, 300, "Change Report Form Received"
  EditBox 60, 5, 40, 15, MAXIS_case_number
  EditBox 160, 5, 45, 15, effective_date
  EditBox 320, 5, 45, 15, date_received
  EditBox 50, 35, 315, 15, address_notes
  EditBox 50, 55, 315, 15, household_notes
  EditBox 50, 75, 315, 15, income_notes
  EditBox 50, 95, 315, 15, shelter_notes
  EditBox 110, 115, 250, 15, asset_notes
  EditBox 50, 135, 315, 15, vehicles_notes
  EditBox 50, 155, 315, 15, other_change_notes
  EditBox 60, 185, 305, 15, actions_taken
  EditBox 60, 205, 305, 15, other_notes
  EditBox 70, 225, 295, 15, verifs_requested
  CheckBox 10, 260, 140, 10, "Check here to navigate to DAIL/WRIT", tikl_nav_check
  DropListBox 270, 255, 95, 20, "Select One:"+chr(9)+"will continue next month"+chr(9)+"will not continue next month", changes_continue
  EditBox 75, 275, 85, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 260, 275, 50, 15
    CancelButton 315, 275, 50, 15
  Text 5, 10, 50, 10, "Case Number:"
  Text 110, 10, 50, 10, "Effective Date:"
  Text 210, 10, 110, 10, "Date Change Reported/Received:"
  GroupBox 5, 25, 365, 150, "Changes Reported:"
  Text 15, 40, 30, 10, "Address:"
  Text 15, 60, 35, 10, "HH Comp:"
  Text 15, 80, 30, 10, "Income:"
  Text 15, 100, 25, 10, "Shelter:"
  Text 15, 120, 95, 10, "Assets (savings or property):"
  Text 15, 140, 30, 10, "Vehicles:"
  Text 15, 160, 20, 10, "Other:"
  Text 10, 190, 45, 10, "Action Taken:"
  Text 10, 210, 45, 10, "Other Notes:"
  Text 10, 230, 60, 10, "Verifs Requested:"
  Text 10, 280, 60, 10, "Worker Signature:"
  Text 180, 260, 90, 10, "The changes client reports:"
  GroupBox 5, 175, 365, 75, "Actions"
EndDialog


' Updated dialog to include field validation
DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation ' Asks user to confirm if they want to cancel

		' Validate that MAXIS number is numeric and less than 8 digits long
		CALL validate_MAXIS_case_number(err_msg, "* ")
		' Validate that worker signature is not blank.
		IF trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		' Validate that worker selects option from dropdown list as to how long change will last
		If changes_continue = "Select One:" THEN err_msg = err_msg & vbNewLine & "* You must select an option from the dropdown list indicating whether the changes reported by the client will continue next month or will not continue next month."
		' Validate that Date Effective field is not empty and is in a proper date format
		If IsDate(trim(effective_date)) = False OR Len(trim(effective_date)) <> 10 Then err_msg = err_msg & vbNewLine & "* The Date Effective field cannot be blank and must be in the MM/DD/YYYY format."
		' Validate that Date Change Reported/Received field is not empty and is in a proper date format
		If IsDate(trim(date_received)) = False OR Len(trim(date_received)) <> 10 Then err_msg = err_msg & vbNewLine & "* The Date Change Reported/Received field cannot be blank and must be in the MM/DD/YYYY format."
		' Validate the change(s) reported fields to ensure that at least one field is filled in
		If address_notes = "" AND household_notes = "" AND asset_notes = "" AND vehicles_notes = "" AND income_notes = "" AND shelter_notes = "" AND other_change_notes = "" THEN err_msg = err_msg & vbNewLine & "* All of the Changes Reported fields are blank. You must enter information in at least one field."
		' Validate the change(s) reported fields to ensure that at least one field is filled in
		If actions_taken = "" AND other_notes = "" AND verifs_requested = "" THEN err_msg = err_msg & vbNewLine & "* All of the Actions fields are blank. You must enter information in at least one field."
		If err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)
LOOP UNTIL are_we_passworded_out = false



' Check if case is privileged and end script if it is privileged
Call navigate_to_MAXIS_screen_review_PRIV("CASE", "NOTE", is_this_priv)
If is_this_priv = True then script_end_procedure("This case is privileged and cannot be accessed. The script will now stop.")
' Confirm that the case is in county and end script if the case is out of county
EMReadScreen county_code, 4, 21, 14
If county_code <> worker_county_code then script_end_procedure("This case is out-of-county, and cannot access CASE:NOTE. The script will now stop.")

'THE CASENOTE----------------------------------------------------------------------------------------------------
'Navigates to case note
Call start_a_blank_case_note
CALL write_variable_in_case_note ("--CHANGE REPORTED-- " & "Date Effective: " & effective_date)
CALL write_bullet_and_variable_in_case_note("Date Received", date_received)
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
CALL write_bullet_and_variable_in_case_note("The changes client reports", changes_continue)
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)

'If we checked to TIKL out, it goes to TIKL and sends a TIKL
IF tikl_nav_check = 1 THEN
	CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
	CALL create_MAXIS_friendly_date(date, 10, 5, 18)
	EMSetCursor 9, 3
END IF
script_end_procedure ("Success! The case note has been created. Send verifications to ECF or case note how the information was received.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------05/01/2023
'--Tab orders reviewed & confirmed----------------------------------------------05/01/2023
'--Mandatory fields all present & Reviewed--------------------------------------05/01/2023
'--All variables in dialog match mandatory fields-------------------------------05/01/2023
'Review dialog names for content and content fit in dialog----------------------05/01/2023
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------05/02/2023
'--CASE:NOTE Header doesn't look funky------------------------------------------05/02/2023
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------05/02/2023 --- When DAIL/WRIT selected, CASE:NOTE is automatically entered and submitted
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used-05/02/2023 -----------------------------------
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------05/02/2023
'--MAXIS_background_check reviewed (if applicable)------------------------------05/02/2023 - N/A
'--PRIV Case handling reviewed -------------------------------------------------05/04/2023
'--Out-of-County handling reviewed----------------------------------------------05/04/2023
'--script_end_procedures (w/ or w/o error messaging)----------------------------05/02/2023
'--BULK - review output of statistics and run time/count (if applicable)--------05/02/2023 - N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------05/02/2023
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------05/02/2023
'--Incrementors reviewed (if necessary)-----------------------------------------05/02/2023
'--Denomination reviewed -------------------------------------------------------05/02/2023
'--Script name reviewed---------------------------------------------------------05/02/2023
'--BULK - remove 1 incrementor at end of script reviewed------------------------05/02/2023 - N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------05/02/2023
'--comment Code-----------------------------------------------------------------05/02/2023
'--Update Changelog for release/update------------------------------------------05/02/2023
'--Remove testing message boxes-------------------------------------------------05/02/2023
'--Remove testing code/unnecessary code-----------------------------------------05/04/2023
'--Review/update SharePoint instructions----------------------------------------05/02/2023
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------05/04/2023
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------To Be Completed
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------To Be Completed
'--Complete misc. documentation (if applicable)---------------------------------05/04/2023
'--Update project team/issue contact (if applicable)----------------------------05/04/2023

