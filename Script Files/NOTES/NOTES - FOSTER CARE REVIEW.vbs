name_of_script = "NOTES - FOSTER CARE REVIEW.vbs"
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

'Dialog---------------------------------------------------------------------------------------------------------------------------
BeginDialog FC_HC_REVIEW, 0, 0, 256, 250, "FOSTER CARE HC REVIEW"
  EditBox 65, 5, 65, 15, Case_number
  EditBox 65, 25, 65, 15, Received
  EditBox 65, 45, 65, 15, Completed_By
  EditBox 130, 70, 105, 15, Social_Worker_or_Probation_Officer
  EditBox 105, 90, 85, 15, Extended_Foster_Care_Date
  EditBox 40, 125, 55, 15, Income
  EditBox 40, 145, 70, 15, Results
  EditBox 75, 190, 110, 15, Worker_Signature
  ButtonGroup ButtonPressed
    OkButton 125, 230, 50, 15
    CancelButton 190, 230, 50, 15
  Text 5, 5, 55, 10, "Case number:"
  Text 5, 25, 45, 10, "Received: "
  Text 5, 45, 60, 10, "Completed By: "
  Text 5, 70, 120, 10, "Social Worker or Probation Officer:"
  Text 5, 95, 95, 15, "Extended Foster Care Date: "
  Text 5, 125, 35, 15, "Income: "
  Text 5, 150, 30, 15, "Results:"
  Text 5, 190, 65, 10, "Worker Signature: "
EndDialog



'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------
'connecting to BlueZone, and grabbing the case number
EMConnect ""
Call MAXIS_case_number_finder(case_number)

'calling the dialog---------------------------------------------------------------------------------------------------------------
DO
	Dialog County_Burial_Application_Received
	IF buttonpressed = 0 THEN stopscript
	IF case_number = "" THEN MsgBox "You must have a case number to continue!"
	IF worker_signature = "" THEN MsgBox "You must enter a worker signature."
LOOP until case_number <> "" and worker_signature <> ""

'checking for an active MAXIS session
CALL check_for_MAXIS(FALSE)

'The case note---------------------------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
CALL write_variable_in_CASE_NOTE("***Foster Care HC REVIEW***")
CALL write_bullet_and_variable_in_CASE_NOTE("Received", Received)
CALL write_bullet_and_variable_in_CASE_NOTE("Completed By", Completed_By)
CALL write_bullet_and_variable_in_CASE_NOTE("Social Worker or Probation Officer", Social_Worker_or_Probation_Officer)
CALL write_bullet_and_variable_in_CASE_NOTE("Extended Foster Care Date", Extended_Foster_Care_Date)
CALL write_bullet_and_variable_in_CASE_NOTE("Income", Income)
CALL write_bullet_and_variable_in_CASE_NOTE("Results", Results)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)
Script_end_procedure("")

