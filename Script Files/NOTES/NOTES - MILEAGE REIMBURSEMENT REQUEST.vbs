'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - MILEAGE REIMBURSEMENT REQUEST.vbs"
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

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog mileage_dialog, 0, 0, 306, 125, "Mileage Reimbursement"
  EditBox 75, 5, 70, 15, case_number
  EditBox 230, 5, 70, 15, date_docs_recd
  EditBox 50, 25, 70, 15, total_reimbursement
  EditBox 230, 25, 70, 15, date_to_accounting
  EditBox 50, 45, 250, 15, docs_reqd
  EditBox 50, 65, 250, 15, other_notes
  EditBox 55, 85, 245, 15, actions_taken
  EditBox 70, 105, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 195, 105, 50, 15
    CancelButton 250, 105, 50, 15
  Text 5, 10, 70, 10, "MAXIS case number:"
  Text 170, 10, 55, 10, "Date Received:"
  Text 5, 30, 45, 10, "Total Amount:"
  Text 165, 30, 60, 10, "Date Sent to Acct:"
  Text 5, 50, 40, 10, "Doc's req'd:"
  Text 5, 70, 45, 10, "Other notes:"
  Text 5, 90, 50, 10, "Actions taken:"
  Text 5, 110, 60, 10, "Worker signature:"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""
'Finds the case number
call MAXIS_case_number_finder(case_number)


'Displays the dialog and navigates to case note
Do
	Dialog Mileage_dialog
	cancel_confirmation
	If case_number = "" then MsgBox "You must have a case number to continue!"
Loop until case_number <> ""

'checking for an active MAXIS session
Call check_for_MAXIS(False)

'the CASE NOTE----------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE(">>>>>MILEAGE REIMBURSEMENT REQUEST - ACTIONS TAKEN<<<<<")
call write_bullet_and_variable_in_case_note("Date received", date_docs_recd)
call write_bullet_and_variable_in_case_note("Total Amount", "$" & total_reimbursement)
call write_bullet_and_variable_in_case_note("Date Sent to Accounting", date_to_accounting)
call write_bullet_and_variable_in_case_note("Docs requested", docs_reqd)
call write_bullet_and_variable_in_case_note("Other notes", other_notes)
call write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
If worker_county_code = "x179" then call write_variable_in_CASE_NOTE("* Please note: DO NOT SCAN!! Accounting will scan into OnBase when processed.")	'Should only do this for Wabasha County, unless other counties request it.
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("")
