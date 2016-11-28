'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - GOOD CAUSE CLAIMED.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 90           'manual run time in seconds
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
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

BeginDialog Good_Cause_Claimed_Dialog, 0, 0, 251, 310, "Child Support Good Cause Claimed"
  EditBox 180, 5, 65, 15, MAXIS_case_number
  DropListBox 135, 30, 105, 15, "Select One:"+chr(9)+"New Claim"+chr(9)+"Annual Redetermination", Claim_Type_droplist
  EditBox 60, 60, 65, 15, Claim_Date
  EditBox 175, 60, 65, 15, Expiration_Date
  EditBox 150, 90, 65, 15, Date_DHS_Claim_Docs
  EditBox 150, 115, 65, 15, Date_DHS_Exp_Docs
  EditBox 45, 140, 195, 15, List_programs
  CheckBox 5, 165, 160, 15, "Supporting documentation has been provided.", Docs_provided_check
  EditBox 180, 185, 65, 15, Supporting_doc_date
  EditBox 180, 210, 65, 15, GC_Review_Date
  EditBox 30, 235, 210, 15, Other_comments
  EditBox 70, 260, 75, 15, Worker_signature
  ButtonGroup ButtonPressed
    OkButton 135, 285, 50, 15
    CancelButton 190, 285, 50, 15
  Text 125, 10, 50, 15, "Case Number"
  Text 5, 30, 130, 15, "Is this a new claim or redetermination?"
  GroupBox 5, 50, 250, 35, "Date Good Cause"
  Text 30, 65, 30, 15, "Claimed"
  Text 135, 65, 35, 15, "Expiration"
  Text 5, 90, 135, 15, "Date DHS-3627, DHS-3632, and DHS-3979 were sent:"
  Text 5, 115, 135, 15, "Date DHS-3630 and DHS-3631 were sent:"
  Text 5, 145, 40, 15, "Programs:"
  Text 5, 185, 175, 15, "Deadline given to provide supporting documentation:"
  Text 5, 205, 165, 20, "Date Good Cause claim will be reviewed (no more than 20 days from present):"
  Text 5, 235, 20, 15, "Other:"
  Text 5, 260, 60, 15, "Worker Signature"
EndDialog

'Script----------------------------------------------
'Connect to Bluezone
EMConnect ""

'Inserts Maxis Case number
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Shows dialog
DO
	DO
		DO
			Dialog Good_Cause_Claimed_Dialog
			IF ButtonPressed = 0 THEN StopScript
			IF IsNumeric(MAXIS_case_number) = FALSE THEN MsgBox "You must type a valid numeric case number."
		LOOP UNTIL IsNumeric(MAXIS_case_number) = TRUE
		IF Claim_Type_droplist = "Select One:" THEN MsgBox "You must select New Claim or Redetermination."
	LOOP UNTIL Claim_Type_droplist <> "Select One:"
	IF worker_signature = "" THEN MsgBox "You must sign your case note!"
LOOP UNTIL worker_signature <> ""

'Checks Maxis for password prompt
CALL check_for_MAXIS(True)

'Navigates to case note
CALL navigate_to_MAXIS_screen("CASE", "NOTE")

'Sends a PF9
PF9

'Writes the case note
CALL write_variable_in_case_note("Child Support Good Cause Exemption " & Claim_Type_droplist)
CALL write_bullet_and_variable_in_case_note("Good Cause claimed on", Claim_date)
CALL write_bullet_and_variable_in_case_note("Good Cause expiration", Expiration_date)
CALL write_bullet_and_variable_in_case_note("DHS-3627 and DHS-3979 were sent on", Date_DHS_Claim_Docs)
CALL write_bullet_and_variable_in_case_note("DHS-3630 and DHS-3631 were sent on", Date_DHS_Exp_Docs)
CALL write_bullet_and_variable_in_case_note("Programs", List_programs)
IF Docs_Provided_Check = 1 THEN CALL write_variable_in_case_note("* Supporting documentation has been provided.")
CALL write_bullet_and_variable_in_case_note("Deadline given to provide supporting documentation ", Supporting_doc_date)
CALL write_bullet_and_variable_in_case_note("Date Good Cause claim will be reviewed", GC_review_date)
CALL write_bullet_and_variable_in_case_note("Other", Other_comments)
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)


CALL script_end_procedure("")
