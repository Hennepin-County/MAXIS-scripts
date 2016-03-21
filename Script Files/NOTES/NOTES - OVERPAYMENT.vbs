'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - OVERPAYMENT.vbs"
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
STATS_manualtime = 120          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'SECTION 02: DIALOGS
BeginDialog overpayment_dialog, 0, 0, 266, 260, "Overpayment dialog"
  EditBox 60, 5, 70, 15, case_number
  EditBox 120, 25, 140, 15, programs_cited
  EditBox 100, 45, 160, 15, Claim_number
  EditBox 120, 65, 140, 15, months_of_overpayment
  EditBox 65, 85, 60, 15, discovery_date
  EditBox 200, 85, 60, 15, established_date
  EditBox 100, 105, 160, 15, reason_for_OP
  EditBox 150, 125, 110, 15, reason_to_be_reported
  EditBox 85, 145, 175, 15, supporting_docs
  EditBox 125, 165, 135, 15, responsible_parties
  EditBox 60, 185, 200, 15, total_amt_of_OP
  EditBox 70, 205, 50, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 80, 240, 50, 15
    CancelButton 135, 240, 50, 15
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 30, 115, 10, "Program(s) overpayment cited for:"
  Text 5, 70, 110, 10, "Month(s)/Year(s) of overpayment:"
  Text 5, 90, 55, 10, "Discovery date:"
  Text 135, 90, 60, 10, "Established date:"
  Text 5, 110, 95, 10, "Reason for OP (Be Specific):"
  Text 5, 150, 80, 10, "Supporting docs/verifs:"
  Text 5, 170, 120, 10, "Responsible parties listed by name:"
  Text 5, 190, 55, 10, "Total amt of OP:"
  Text 5, 210, 65, 10, "Sign the case note:"
  Text 5, 50, 95, 10, "Claim Number(s) if available: "
  Text 5, 130, 140, 10, "When/why should this have been reported: "
EndDialog

'SECTION 03: THE SCRIPT----------------------------------------------------------------------------------------------------
'connecting to MAXIS
EMConnect ""
'grabbing case number
Call MAXIS_case_number_finder(case_number)

DO
	Do
		Dialog overpayment_dialog
		cancel_confirmation
		If case_number = "" then MsgBox "You must have a case number to continue!"
	Loop until case_number <> ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false


'checking for an active MAXIS session
Call check_for_MAXIS(False)

If Claim_number = "" Then
	Claim_number = "Not available at this time"
end if

'THE CASE NOTE----------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
call write_variable_in_CASE_NOTE("**OVERPAYMENT/CLAIM ESTABLISHED**")
call write_bullet_and_variable_in_case_note("Program(s) overpayment cited for", programs_cited) 
call write_bullet_and_variable_in_case_note("Claim Number(s)", Claim_number) 
call write_bullet_and_variable_in_case_note("Month(s) of overpayment", months_of_overpayment) 
call write_bullet_and_variable_in_case_note("Discovery date", discovery_date) 
call write_bullet_and_variable_in_case_note("Established date", established_date) 
call write_bullet_and_variable_in_case_note("Reason for overpayment", reason_for_OP) 
call write_bullet_and_variable_in_case_note("When/Why should this have been reported", reason_to_be_reported) 
call write_bullet_and_variable_in_case_note("Supporting documents/verifications", supporting_docs) 
call write_bullet_and_variable_in_case_note("Responsible parties", responsible_parties) 
call write_bullet_and_variable_in_case_note("Total overpayment amount", total_amt_of_OP) 
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("")
