'STATS GATHERING----------------------------------------------------------------------------------------------------
'TO DO - update CLOS accordingly
name_of_script = "NOTES - DECEASED RESIDENT SUMMARY.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 150                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
'END OF stats block========================================================================================================

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
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THIS SCRIPT IS BEING USED IN A WORKFLOW SO DIALOGS ARE NOT NAMED
'DIALOGS MAY NOT BE DEFINED AT THE BEGINNING OF THE SCRIPT BUT WITHIN THE SCRIPT FILE

'There is only one dialog in this script and so it can be defined in the beginning, but still is unnamed
'THE SCRIPT

'Connects to BlueZone
EMConnect ""
'Calls a MAXIS case number
call MAXIS_case_number_finder(MAXIS_case_number)

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
'TO DO - change client to name here?
'TO DO - reformat dialog once validation and logic finalized
BeginDialog Dialog1, 0, 0, 236, 240, "Deceased Client Summary"
  Text 5, 10, 50, 10, "Case Number"
  EditBox 60, 5, 65, 15, MAXIS_case_number
  Text 5, 35, 50, 10, "Date of Death"
  EditBox 60, 30, 65, 15, date_of_death
  Text 5, 55, 50, 10, "Place of Death"
  EditBox 60, 50, 170, 15, place_of_death
  Text 5, 75, 65, 10, "Surviving Spouse?"
  CheckBox 110, 75, 55, 10, "(check if yes)", surviving_spouse_checkbox
  Text 5, 90, 55, 10, "MA Lien on File?"
  CheckBox 110, 90, 55, 10, "(check if yes)", MA_lien_on_file_checkbox
  Text 5, 105, 100, 10, "Is servicing county also CFR?"
  CheckBox 110, 105, 55, 10, "(check if yes)", servicing_county_checkbox
  Text 5, 120, 100, 10, "Transfer case to CFR?"
  CheckBox 110, 120, 55, 10, "(check if yes)", transfer_to_CFR_checkbox
  Text 5, 135, 80, 20, "Refer file for possible estate collection?"
  CheckBox 110, 140, 55, 10, "(check if yes)", collection_checkbox
  Text 5, 165, 35, 10, "Other info"
  EditBox 65, 160, 165, 15, other_info
  Text 5, 185, 45, 10, "Actions taken"
  EditBox 65, 180, 165, 15, actions_taken
  Text 5, 205, 60, 10, "Worker Signature"
  EditBox 65, 200, 95, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 115, 220, 40, 15
    CancelButton 160, 220, 40, 15
  GroupBox 145, 0, 85, 45, "Links"
  ButtonGroup ButtonPressed
    PushButton 150, 10, 75, 15, "HSR Manual", Button4
    PushButton 150, 25, 75, 15, "Print Document/Case", Button6
EndDialog
'TO DO - keeping previous dialog for reference, delete once finalized
' BeginDialog Dialog1, 0, 0, 206, 190, "Deceased Client Summary"
'   Text 5, 10, 50, 10, "Case Number"
'   EditBox 65, 5, 50, 15, MAXIS_case_number
'   Text 5, 30, 50, 10, "Date of Death"
'   EditBox 65, 25, 50, 15, date_of_death
'   Text 5, 50, 50, 10, "Place of Death"
'   EditBox 65, 45, 100, 15, place_of_death
'   Text 5, 65, 65, 10, "Surviving Spouse?"
'   CheckBox 105, 65, 55, 10, "(check if yes)", surviving_spouse_checkbox
'   Text 5, 80, 60, 10, "MA Lien on File?"
'   CheckBox 105, 80, 55, 10, "(check if yes)", MA_lien_on_file_checkbox
'   CheckBox 5, 100, 110, 10, "Is servicing county also CFR?", servicing_county_checkbox
'   CheckBox 120, 100, 85, 10, "Transfer case to CFR?", transfer_to_CFR_checkbox
'   CheckBox 5, 115, 145, 10, "Refer file for possible estate collection?", collection_checkbox
'   Text 5, 130, 35, 10, "Other info"
'   EditBox 65, 130, 135, 15, other_info
'   Text 5, 150, 45, 10, "Actions taken"
'   EditBox 65, 150, 135, 15, actions_taken
'   Text 5, 175, 60, 10, "Worker Signature"
'   EditBox 65, 170, 50, 15, worker_signature
'   ButtonGroup ButtonPressed
' 	OkButton 120, 170, 40, 15
' 	CancelButton 160, 170, 40, 15
' EndDialog
'Do loop for Deceased Client Summary Shows dialog and creates and displays an error message if worker completes things incorrectly.
DO
     DO
    	err_msg = ""
    	DIALOG Dialog1					'Calling a dialog without a assigned variable will call the most recently defined dialog
    	cancel_confirmation
    	'case number required for case note
    	IF isnumeric  (MAXIS_case_number) = false THEN err_msg = err_msg & "Please enter a case number." & VBnewline
    	'valid date required
    	IF isDate (date_of_death)=false then err_msg=err_msg & "Please enter a valid date." & VBNewline
    	'worker signature required
    	IF worker_signature = "" THEN err_msg = err_msg & "Please enter your worker signature." & VBNewline
    LOOP UNTIL err_msg = ""
    call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'Checks MAXIS for password prompt
CALL check_for_MAXIS(false)

'Navigates to case note
start_a_blank_CASE_NOTE
'writes case note for deceased client summary
'TO DO - change client to name here?
CALL write_variable_in_Case_Note("--Deceased Client Summary--")
CALL write_bullet_and_variable_in_Case_Note("Date of Death", date_of_death)
CALL write_bullet_and_variable_in_Case_Note("Place of Death", place_of_death)
IF surviving_spouse_Checkbox = 1 THEN CALL write_variable_in_Case_Note("* There is a surviving spouse.")
IF MA_lien_on_file_checkbox = 1 THEN CALL write_variable_in_Case_Note("* MA lien on file.")
IF servicing_county_Checkbox = 1 THEN CALL write_variable_in_Case_Note("* Servicing county is also CFR.")
IF transfer_to_CFR_Checkbox = 1 THEN CALL write_variable_in_Case_Note("* Transfer case to CFR.")
IF collection_Checkbox = 1 THEN CALL write_variable_in_Case_Note("* Refer file for possible estate collection.")
CALL write_bullet_and_variable_in_Case_Note("Other Info", other_info)
CALL write_bullet_and_variable_in_Case_Note("Actions Taken", actions_taken)
CALL write_variable_in_Case_Note("---")
CALL write_variable_in_Case_Note(worker_signature)

'transmit to save case note
PF3

'Script ends & reminds worker to update STAT MEMB if need be
script_end_procedure("Success! Case note has been added. Make sure date of death is entered on STAT MEMB and proceed as needed.")

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

