'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - MNSURE RETRO HC APPLICATION.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 180          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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
call changelog_update("01/12/2017", "Changed verbiage from Curam to METS.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE Dialog--------------------------------------------------------------------------------------------------------------------
BeginDialog MNSure_HC_Appl_dialog, 0, 0, 326, 265, "MNSure Retro HC Application"
  EditBox 85, 25, 55, 15, MAXIS_case_number
EditBox 240, 25, 75, 15, METS_case_number
  EditBox 120, 50, 60, 15, HC_Appl_date_Recvd
  EditBox 235, 50, 75, 15, time_gap_between
  DropListBox 135, 70, 60, 15, "Select One..."+chr(9)+"1 Month"+chr(9)+"2 Months"+chr(9)+"3 Months", retro_coverage_months
  EditBox 165, 90, 65, 15, hc_closed_120days
  EditBox 125, 110, 185, 15, HH_members_requesting
  DropListBox 90, 130, 60, 15, "Select One..."+chr(9)+"Approved"+chr(9)+"Denied"+chr(9)+"Pending", HC_Appl_status
  EditBox 85, 150, 230, 15, missing_documents
  EditBox 95, 170, 220, 15, other_notes
  EditBox 60, 190, 255, 15, action_done_taken
  CheckBox 10, 215, 125, 10, "Navigate to TIKL for 10 Day Return", tikl_checkbox
  CheckBox 165, 215, 75, 10, "Updated in MMIS", mmis_checkbox
  EditBox 85, 240, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 210, 240, 50, 15
    CancelButton 265, 240, 50, 15
Text 165, 30, 70, 10, "METS Case Number:"
  Text 10, 55, 105, 10, "MNSure Application Rec'd Date:"
  Text 10, 115, 115, 10, "Retro HC Coverage Requested for:"
  Text 10, 135, 75, 10, "HC Application Status:"
  Text 10, 30, 70, 10, "Maxis Case Number:"
  Text 10, 95, 150, 10, " If HC closed within 120 days, date of closure:"
  Text 195, 55, 40, 10, "Gap Month:"
  Text 10, 155, 70, 10, "Verifications Needed:"
  Text 10, 75, 120, 10, "Number of Retro Months Requested:"
  Text 10, 195, 50, 10, "Action Taken:"
  Text 10, 175, 80, 10, "Other Comments/Notes:"
Text 10, 10, 220, 10, "MA application in METS, Retro eligibility determination requested "
  Text 10, 245, 75, 10, "Sign Your Case Note:"
EndDialog

'THE Script-----------------------------------------------------------------------------------------------------------------------------

'Connects to Bluzone
EMConnect ""

'Brings Bluezone to the front
EMFocus

'Grabs the MAXIS Case Number
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Shows Dialog

DO
	DO
		DO
			DO
				Dialog MNSure_HC_Appl_dialog
				cancel_confirmation
				IF worker_signature = "" THEN MsgBox "You must sign your case note!"
				IF IsNumeric(MAXIS_case_number) = FALSE THEN MsgBox "You must type a valid numeric case number!"
				IF retro_coverage_months = "Select One..." THEN MsgBox "Please select how many retro months are requested!"
				IF len(METS_case_number)<> 8 THEN MsgBox "Please enter an 8 digit METS case number!"
			Loop until len(METS_case_number) = 8
		Loop until retro_coverage_months <> "Select One..."
		IF HH_members_requesting = "" THEN MsgBox "Enter HH members requesting retro HC coverage!"
	Loop until HH_members_requesting <> ""
	IF HC_Appl_status = "Select One..." THEN MsgBox "Select a status for this application!"
Loop until HC_Appl_status <> "Select One..."


'Opens a New Case Note
call start_a_blank_CASE_NOTE

'Writes the Case Note
CALL write_variable_in_CASE_NOTE("---MNSure Retro HC Application---")
Call write_bullet_and_variable_in_CASE_NOTE("METS Case Number", METS_case_number)
Call write_bullet_and_variable_in_CASE_NOTE("MNSure Application Rec'd", HC_Appl_date_Recvd)
Call write_bullet_and_variable_in_CASE_NOTE("Gap Month", time_gap_between)
Call write_bullet_and_variable_in_CASE_NOTE("Number of Retro Months Requested", retro_coverage_months)
Call write_bullet_and_variable_in_CASE_NOTE("HC Closed within 120 Days On", hc_closed_120days)
Call write_bullet_and_variable_in_CASE_NOTE("Retro HC Coverage Requested For", HH_members_requesting)
Call write_bullet_and_variable_in_CASE_NOTE("Retro HC Application Status", HC_Appl_status)
Call write_bullet_and_variable_in_CASE_NOTE("Verifications Needed", missing_documents)
CALL write_bullet_and_variable_in_case_note("Other Comments/Notes", other_notes)
Call write_bullet_and_variable_in_CASE_NOTE("Actions taken", action_done_taken)
IF mmis_checkbox = checked THEN CALL write_variable_in_case_note("* Updated in MMIS")
IF tikl_checkbox = checked THEN CALL write_variable_in_case_note("* TIKL'd for a 10 day return")
CALL write_variable_in_case_note("---")
Call write_variable_in_CASE_NOTE(worker_signature)

'The TIKL
IF tikl_checkbox = 1 THEN
	CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
	CALL create_maxis_friendly_date(date, 10, 5, 18)
	EMSetCursor 9, 3
END IF

script_end_procedure ("")
