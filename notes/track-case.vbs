'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - Track Case.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 10                      'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
'END OF stats block=========================================================================================================

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

call changelog_update("10/14/2025", "Initial version.", "Dave Courtright, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'THIS SCRIPT IS NOT BEING ACCESSED FROM GITHUB - there is a LOCAL FILE - use that for updates/action.'
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

BeginDialog Dialog1, 0, 0, 216, 120, "Postponed Approval Required"
  EditBox 65, 50, 55, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    OkButton 115, 95, 40, 15
    CancelButton 170, 95, 40, 15
  Text 10, 55, 50, 10, "Case number:"
  GroupBox 5, 5, 220, 40, "When to use this script:"
  Text 10, 15, 195, 25, "Use this script if SNAP or MFIP results are ready for approval but cannot be approved due to 2025 Federal Government Shutdown."
  CheckBox 120, 75, 50, 10, "SNAP", snap_check
  CheckBox 175, 75, 50, 10, "MFIP", MFIP_check
  Text 10, 75, 100, 10, "Program(s) needing Approval:"
EndDialog


Do
	err_msg = ""
    Dialog Dialog1
	Cancel_without_confirmation
    Call validate_MAXIS_case_number(err_msg, "*")
    If MFIP_check = False and snap_check = False Then
        err_msg = err_msg & "You must select at least one program needing approval."
    End If
    IF MAXIS_case_number = "" THEN
        err_msg = err_msg & "You must enter a case number." & vbNewLine
    End IF
    IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
LOOP UNTIL err_msg = ""

IF snap_check = 1 THEN
    program_list = "~SNAP~"
End IF
IF MFIP_check = 1 THEN
    IF program_list <> "" THEN
        program_list = program_list & " and ~MFIP~"
    ELSE
        program_list = "~MFIP~"
    End IF
End IF
start_a_blank_CASE_NOTE
CALL write_variable_in_CASE_NOTE(program_list & " approval postponed due to 2025 Federal Government Shutdown.")
Call write_variable_in_CASE_NOTE(program_list & " ready to approve, but benefits cannot be approved at this time due to expiration of funding due to federal government shutdown. Case requires future follow up for review and approval.")
Call write_variable_in_CASE_NOTE("Case added to tracking list.")
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)


script_end_procedure("Success! Your case number, " & MAXIS_case_number & ", has been captured for future follow up. Programs needing approval: " & program_list & ".")

