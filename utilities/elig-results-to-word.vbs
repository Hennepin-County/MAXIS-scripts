'STATS GATHERING--------------------------------------------------------------------------------------------------------------
name_of_script = "UTILITIES - ELIG RESULTS TO WORD.vbs"
start_time = timer

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
		FuncLib_URL = "C:\MAXIS-scripts\MASTER FUNCTIONS LIBRARY.vbs"
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
call changelog_update("06/02/2021", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


'TODO - there is currently no handling for HC
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 231, 195, "Case Number to Read ELIG Results"
  EditBox 65, 65, 50, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    OkButton 115, 170, 50, 15
    CancelButton 170, 170, 50, 15
  Text 15, 15, 175, 20, "This script works by pulling all of the information from a specific version of ELIG into word."
  Text 15, 40, 195, 20, "Enter the Case Number and Navigate in MAXIS now to the start of the ELIG version that you would like copied."
  Text 15, 70, 45, 10, "Case Number"
  GroupBox 15, 90, 205, 40, "NAVIGATE IN MAXIS NOW TO THE ELIG VERSION TO COPY"
  Text 30, 105, 105, 10, "- ELIG can be for any program."
  Text 30, 115, 165, 10, "- Version of ELIG can be approved or unapproved."
  Text 15, 145, 180, 20, "When the script continues, it will look for the first page of ELIG Results. If not found, the dialog will reappear."
EndDialog

Do
	Do
		Do
			err_msg = ""

			dialog Dialog1
			cancel_without_confirmation

			Call validate_MAXIS_case_number(err_msg, "*")

		Loop until err_msg = ""
		elig_results_program_found = ""
		EMReadScreen MX_line_3, 78, 3, 2
		If InStr(MX_line_3, "DWPR") Then elig_results_program_found = "DWP"
		If InStr(MX_line_3, "MFPR") Then elig_results_program_found = "MFIP"
		If InStr(MX_line_3, "MSPR") Then elig_results_program_found = "MSA"
		If InStr(MX_line_3, "GAPR") Then elig_results_program_found = "GA"
		If InStr(MX_line_3, "CAPR") Then elig_results_program_found = "Cash Denial"
		If InStr(MX_line_3, "GRPR") Then elig_results_program_found = "GRH"
		If InStr(MX_line_3, "FCSM") Then elig_results_program_found = "IV-E"
		If InStr(MX_line_3, "EMPR") Then elig_results_program_found = "EMER"
		If InStr(MX_line_3, "FSPR") Then elig_results_program_found = "SNAP"
		If elig_results_program_found = "" Then MsgBox "MAXIS must be at the first page of Eligibility Results for this script to run." & vbCr & vbCr & "The dialog will return." & vbCr & vbCr & "NAVIGATE TO ELIG RESULTS WHILE THE DIALOG IS UP."

		EMReadScreen version_number, 2, 2, 12
	Loop until elig_results_program_found <> ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = False

'Now we make sure the check for password transmit wasn't a problem.
'Need to navigate to the first page based on program. Each program has different coordinates

If elig_results_program_found = "SNAP" Then Call write_value_and_transmit("FSPR", 19, 70)

'confirm we are at the right version

'Read each line of the results.
'DO WE NEED TO OPEN POP-UPs?