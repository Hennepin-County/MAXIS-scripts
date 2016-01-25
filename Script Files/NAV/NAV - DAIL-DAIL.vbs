'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NAV - DAIL-DAIL.vbs"
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
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 10                      'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
'END OF stats block=========================================================================================================

'DIALOGS-----------------------------------------------------------------------------------
BeginDialog worker_dialog, 0, 0, 171, 45, "Worker dialog"
  Text 5, 10, 130, 10, "Enter the worker number (last 3 digits):"
  EditBox 135, 5, 30, 15, worker_number
  ButtonGroup ButtonPressed
    OkButton 30, 25, 50, 15
    CancelButton 90, 25, 50, 15
EndDialog

'THE SCRIPT------------------------------------------------------------------------------------------------------

'Determines if user needs the "select-a-worker" version of this nav script, based on the global variables file.
result = filter(users_using_select_a_user, ucase(windows_user_ID))
IF ubound(result) >= 0 OR all_users_select_a_worker = TRUE THEN
	select_a_worker = TRUE
ELSE
	select_a_worker = FALSE
END IF

'If we have to select a worker, it shows the dialog for it.
IF select_a_worker = TRUE THEN
	Dialog worker_dialog
	IF ButtonPressed = cancel THEN StopScript
END IF

'Determines the county code (a custom function involving multicounty agencies being given a proxy access as a specific county).
call worker_county_code_determination(worker_county_code, two_digit_county_code)

'Connects to BlueZone
EMConnect ""

'This checks to maks sure we're in MAXIS.
call check_for_MAXIS(True)

'Finds a MAXIS case number (if applicable).
call MAXIS_case_number_finder(case_number)

'Navigates to DAIL/DAIL
call navigate_to_MAXIS_screen("DAIL", "DAIL")

'Inputs worker_number variable to the DAIL screen, which only comes up for users with the "select-a-worker" option.
IF worker_number <> "" THEN
	EMWriteScreen worker_county_code & worker_number, 21, 6
	transmit
END IF

script_end_procedure("")  'Ends script
