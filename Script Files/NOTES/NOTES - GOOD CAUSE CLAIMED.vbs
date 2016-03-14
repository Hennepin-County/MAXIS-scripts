'Option Explicit
'DIM beta_agency, url, req, fso, name_of_script, start_time, Funclib_url,run_another_script_fso, fso_command, text_from_the_other_script, run_locally, default_directory

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - GOOD CAUSE CLAIMED.vbs"
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
STATS_manualtime = 90           'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'DIM ButtonGroup_ButtonPressed, ButtonPressed, MAXIS_check, Claim_date, Expiration_date, Date_DHS_Claim_Docs, Date_DHS_Exp_Docs, Docs_provided_check, Good_Cause_Claimed_Dialog, Case_Number, Date_DHS_docs_sent, List_programs, Supporting_doc_date, GC_Review_Date, Other_comments, Worker_signature, Claim_Type_droplist

BeginDialog Good_Cause_Claimed_Dialog, 0, 0, 251, 310, "Child Support Good Cause Claimed"
  EditBox 180, 5, 65, 15, Case_Number
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
CALL MAXIS_case_number_finder(case_number)

'Shows dialog
DO
	DO
		DO
			Dialog Good_Cause_Claimed_Dialog
			IF ButtonPressed = 0 THEN StopScript
			IF IsNumeric(case_number) = FALSE THEN MsgBox "You must type a valid numeric case number."
		LOOP UNTIL IsNumeric(case_number) = TRUE
		IF Claim_Type_droplist = "Select One:" THEN MsgBox "You must select New Claim or Redetermination."
	LOOP UNTIL Claim_Type_droplist <> "Select One:"
	IF worker_signature = "" THEN MsgBox "You must sign your case note!"
LOOP UNTIL worker_signature <> ""

'Checks Maxis for password prompt
CALL check_for_MAXIS(True)

'Navigates to case note
CALL navigate_to_screen("CASE", "NOTE")

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
