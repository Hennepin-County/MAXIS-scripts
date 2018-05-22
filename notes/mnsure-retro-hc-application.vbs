'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - MNSURE RETRO HC APPLICATION.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 230          'manual run time in seconds
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
call changelog_update("09/25/2017", "Enhancements requested by MNSURE team include: Updated navigation and standarization, removed setting a TIKL, and added functionality that sends an email to the specified team with applicable case information.", "MiKayla Handley, Hennepin County")
call changelog_update("08/08/17", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'--------------------------------------------------------------------------------------------------------------------Dialog
BeginDialog MNSure_HC_Appl_dialog, 0, 0, 326, 230, "MNsure Retro HC App, Retro Elig Determination Requested"
  EditBox 85, 5, 75, 15, curam_case_number
  EditBox 85, 25, 75, 15, MAXIS_case_number
  EditBox 85, 45, 75, 15, HH_members_requesting
  DropListBox 260, 5, 60, 15, "Select One..."+chr(9)+"1 Month"+chr(9)+"2 Months"+chr(9)+"3 Months", retro_coverage_months
  DropListBox 260, 25, 60, 15, "Select One..."+chr(9)+"A"+chr(9)+"B"+chr(9)+"C"+chr(9)+"D"+chr(9)+"E", retro_scenario_dropbox
  DropListBox 260, 45, 60, 15, "Select One..."+chr(9)+"YES"+chr(9)+"NO", task_created_dropbox
  EditBox 85, 65, 235, 15, verfs_needed
  EditBox 85, 85, 235, 15, forms_needed
  EditBox 85, 105, 235, 15, other_notes
  EditBox 85, 125, 235, 15, action_taken
  CheckBox 10, 150, 165, 10, "Case Correction and Transfer Use Form Sent (A)", forms_sent_checkbox
  CheckBox 10, 170, 140, 10, "Email sent to HSPH.EWS.Team.603 (B)", EMAIL_603_B_checkbox
  CheckBox 10, 190, 145, 10, "Email sent to HSPH.EWS.Team.601 (E/C)", EMAIL_601_EC_checkbox
  CheckBox 190, 150, 90, 10, "STAT Panels Requested", STAT_added_checkbox
  CheckBox 190, 170, 70, 10, "Updated in MMIS", mmis_checkbox
  EditBox 75, 210, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 215, 210, 50, 15
    CancelButton 270, 210, 50, 15
  Text 170, 10, 85, 10, "Retro Months Requested:"
  Text 10, 130, 45, 10, "Action Taken:"
  Text 10, 110, 60, 10, "Comments/Notes:"
  Text 10, 215, 60, 10, "Worker Signature:"
  Text 205, 50, 45, 10, "Task Created:"
  Text 200, 30, 50, 10, "Retro Scenario:"
  Text 10, 90, 50, 10, "Forms Needed:"
  Text 10, 10, 70, 10, "METS Case Number:"
  Text 10, 30, 70, 10, "Maxis Case Number:"
  Text 10, 50, 70, 10, "Retro Requested for:"
  Text 10, 70, 70, 10, "Verifications Needed:"
EndDialog

'-------------------------------------------------------------------------------------------------------script
EMConnect ""
EMFocus
CALL MAXIS_case_number_finder(MAXIS_case_number)

DO
	DO
		DO
			DO
				Dialog MNSure_HC_Appl_dialog
				cancel_confirmation
				IF worker_signature = "" THEN MsgBox "Please sign your case note."
				IF IsNumeric(MAXIS_case_number) = FALSE THEN MsgBox "You must type a valid numeric case number." 
				IF retro_coverage_months = "Select One..." THEN MsgBox "Please select how many retro months are requested"
				IF len(curam_case_number)<> 8 THEN MsgBox "Please enter an 8 digit Curam case number"
			Loop until len(curam_case_number) = 8
		Loop until retro_coverage_months <> "Select One..."
		IF HH_members_requesting = "" THEN MsgBox "Enter HH members requesting retro HC coverage!"
	Loop until HH_members_requesting <> ""
	IF retro_scenario_dropbox = "Select One..." THEN MsgBox "Select a scenario for this application"
Loop until retro_scenario_dropbox <> "Select One..."

date_due = dateadd("d", +10, date)

'------------------------------------------------------------------------------------Case Note					
call start_a_blank_CASE_NOTE
CALL write_variable_in_CASE_NOTE("---MNsure Retro HC Application, Retro Eligibility Determination Requested---")
Call write_bullet_and_variable_in_CASE_NOTE("METS Case Number", curam_case_number)
Call write_bullet_and_variable_in_CASE_NOTE("MAXIS Case Number", MAXIS_case_number)
Call write_bullet_and_variable_in_CASE_NOTE("Number of Retro Months Requested", retro_coverage_months)
Call write_bullet_and_variable_in_CASE_NOTE("Retro HC Coverage Requested For", HH_members_requesting)
Call write_bullet_and_variable_in_CASE_NOTE("Retro Scenario", retro_scenario_dropbox)
Call write_bullet_and_variable_in_CASE_NOTE("Verifications Requested", verfs_needed)
Call write_bullet_and_variable_in_CASE_NOTE("Forms Needed:", forms_needed)
Call write_bullet_and_variable_in_CASE_NOTE("Due Date:", date_due)
Call write_bullet_and_variable_in_CASE_NOTE("Task Created:", task_created_dropbox)
Call write_bullet_and_variable_in_CASE_NOTE("Action Taken", action_done_taken)
CALL write_bullet_and_variable_in_case_note("Comments/Notes", other_notes)
IF mmis_checkbox = checked THEN CALL write_variable_in_case_note("* Updated MMIS")
IF STAT_added_checkbox = checked THEN CALL write_variable_in_case_note("* STAT Panel(s) To Be Added Sent")
IF forms_sent_checkbox = checked THEN CALL write_variable_in_case_note("* Case Correction and Transfer Use Form Sent")                   
IF EMAIL_603_B_checkbox = checked THEN CALL write_variable_in_case_note("* Other Actions: Email sent to Team 603")   
IF (EMAIL_601_EC_checkbox = checked and retro_scenario_dropbox = "E") THEN CALL write_variable_in_case_note("* Other Actions: Case requires REIN, Email sent to Team 601 for processing")
IF (EMAIL_601_EC_checkbox = checked and retro_scenario_dropbox = "C" or retro_scenario_dropbox ="D") THEN CALL write_variable_in_case_note("* Other Actions: Items listed have been received, Email sent to Team 601 for follow up") 

CALL write_variable_in_case_note("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

'--------------------------------------------------------------------------------------------EMAIL
IF EMAIL_603_B_checkbox = checked THEN CALL create_outlook_email("HSPH.EWS.Team.603", "", "* STAT panel(s) needed, please follow up", "MAXIS case #" & MAXIS_case_number & vbcr & "Case name: " & HH_members_requesting, "", TRUE)
IF (EMAIL_601_EC_checkbox = checked and retro_scenario_dropbox = "E") THEN CALL create_outlook_email("HSPH.EWS.Team.601", "", "* Case requires REIN, please follow up.", "MAXIS case #" & MAXIS_case_number & vbcr & "Case name: " & HH_members_requesting, "", TRUE)
IF (EMAIL_601_EC_checkbox = checked and retro_scenario_dropbox = "C") THEN CALL create_outlook_email("HSPH.EWS.Team.601", "", "* Items listed have been received, please follow up.", "MAXIS case #" & MAXIS_case_number & vbcr & "Case name: " & HH_members_requesting, "", TRUE)

'--------------------------------------------------------------------------------------------TIKL
IF tikl_checkbox = 1 THEN
	CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
	CALL create_maxis_friendly_date(date, 10, 5, 18)
	EMSetCursor 9, 3
END IF	

script_end_procedure ("")