'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - FOSTER CARE HCAPP.vbs"
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
STATS_manualtime = 250         'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'Dialog---------------------------------------------------------------------------------------------------------------------------
BeginDialog Foster_Care_HCAPP, 0, 0, 326, 335, "Foster Care HCAPP "
  EditBox 65, 5, 65, 15, Case_number
  EditBox 105, 25, 65, 15, Date_Received_In_Agency
  EditBox 75, 45, 65, 15, Completed_By
  EditBox 115, 70, 80, 15, Date_of_Agency_Responsiblity
  EditBox 40, 90, 145, 15, IV_E
  EditBox 130, 115, 105, 15, Social_Worker_or_Probation_Officer
  EditBox 40, 135, 160, 15, AREP
  EditBox 55, 155, 70, 15, Income
  EditBox 75, 180, 90, 15, Retro_Requested
  EditBox 45, 200, 130, 15, OHC
  EditBox 95, 235, 160, 15, Verifications_Requested
  EditBox 65, 260, 160, 15, Results
  EditBox 75, 290, 110, 15, Worker_Signature
  ButtonGroup ButtonPressed
    OkButton 215, 320, 50, 15
    CancelButton 275, 320, 50, 15
  Text 5, 5, 55, 10, "Case number:"
  Text 5, 25, 90, 10, "Date Received In Agency: "
  Text 5, 45, 60, 10, "Completed By: "
  Text 5, 70, 100, 10, "Date of Agency Responsiblity: "
  Text 5, 90, 25, 10, "IV E:"
  Text 5, 115, 120, 10, "Social Worker or Probation Officer:"
  Text 5, 135, 35, 10, "AREP: "
  Text 5, 155, 35, 15, "Income: "
  Text 5, 180, 65, 15, "Retro Requested: "
  Text 5, 200, 25, 15, "OHC:"
  Text 5, 235, 85, 15, "Verifications Requested: "
  Text 5, 260, 40, 10, "Results: "
  Text 5, 290, 65, 10, "Worker Signature: "
EndDialog



'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------
'connecting to BlueZone, and grabbing the case number
EMConnect ""
Call MAXIS_case_number_finder(case_number)

'calling the dialog---------------------------------------------------------------------------------------------------------------
DO
	Dialog Foster_Care_HCAPP
	IF buttonpressed = 0 THEN stopscript
	IF case_number = "" THEN MsgBox "You must have a case number to continue!"
	IF worker_signature = "" THEN MsgBox "You must enter a worker signature."
LOOP until case_number <> "" and worker_signature <> ""

'checking for an active MAXIS session
CALL check_for_MAXIS(FALSE)

'The case note---------------------------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
CALL write_variable_in_CASE_NOTE("***Foster Care HCAPP***")
CALL write_bullet_and_variable_in_CASE_NOTE("Date Received In Agency", Date_Received_In_Agency)
CALL write_bullet_and_variable_in_CASE_NOTE("Completed By", Completed_By)
CALL write_bullet_and_variable_in_CASE_NOTE("Date of Agency Responsibility", Date_Of_Agency_Responsiblity)
CALL write_bullet_and_variable_in_CASE_NOTE("IV E", IV_E)
CALL write_bullet_and_variable_in_CASE_NOTE("Social Worker or Probation Officer", Social_Worker_OR_Probation_Officer)
CALL write_bullet_and_variable_in_CASE_NOTE("AREP", Arep)
CALL write_bullet_and_variable_in_CASE_NOTE("Income", Income)
CALL write_bullet_and_variable_in_CASE_NOTE("Retro Requested", Retro_Requested)
CALL write_bullet_and_variable_in_CASE_NOTE("OHC", OHC)
CALL write_bullet_and_variable_in_CASE_NOTE("Verifications Requested", Verifications_Requested)
CALL write_bullet_and_variable_in_CASE_NOTE("Results", Results)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

Script_end_procedure("")
