'STATS GATHERING----------------------------------------------------------------------------------------------------
'IMPORTANT!!! change the name part ..." NOTES - CAF.vbs "...to the file you want this to open.
name_of_script = "NOTES - MNSURE RETRO HC APPLICATION.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
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

'Required for statistical purposes==========================================================================================
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 180          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'THE Dialog--------------------------------------------------------------------------------------------------------------------
BeginDialog MNSure_HC_Appl_dialog, 0, 0, 326, 265, "MNSure Retro HC Application"
  EditBox 85, 25, 55, 15, case_number
  EditBox 240, 25, 75, 15, curam_case_number
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
  Text 165, 30, 70, 10, "Curam Case Number:"
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
  Text 10, 10, 220, 10, "MA application in Curam, Retro eligibility determination requested "
  Text 10, 245, 75, 10, "Sign Your Case Note:"
EndDialog

'THE Script-----------------------------------------------------------------------------------------------------------------------------

'Connects to Bluzone
EMConnect ""

'Brings Bluezone to the front
EMFocus

'Grabs the MAXIS Case Number
CALL MAXIS_case_number_finder(case_number)

'Shows Dialog

DO
	DO
		DO
			DO
				Dialog MNSure_HC_Appl_dialog
				cancel_confirmation
				IF worker_signature = "" THEN MsgBox "You must sign your case note!"
				IF IsNumeric(case_number) = FALSE THEN MsgBox "You must type a valid numeric case number!" 
				IF retro_coverage_months = "Select One..." THEN MsgBox "Please select how many retro months are requested!"
				IF len(curam_case_number)<> 8 THEN MsgBox "Please enter an 8 digit Curam case number!"
			Loop until len(curam_case_number) = 8
		Loop until retro_coverage_months <> "Select One..."
		IF HH_members_requesting = "" THEN MsgBox "Enter HH members requesting retro HC coverage!"
	Loop until HH_members_requesting <> ""
	IF HC_Appl_status = "Select One..." THEN MsgBox "Select a status for this application!"
Loop until HC_Appl_status <> "Select One..."


'Opens a New Case Note					
call start_a_blank_CASE_NOTE

'Writes the Case Note
CALL write_variable_in_CASE_NOTE("---MNSure Retro HC Application---")
Call write_bullet_and_variable_in_CASE_NOTE("Curam Case Number", curam_case_number)
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

