'GATHERING STATS===========================================================================================
name_of_script = "NOTES - SHELTER-HOMELESSNESS VERIFIED.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 180
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================

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
call changelog_update("06/19/2017", "Initial version.", "MiKayla Handley")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'-------------------------------------------------------------------------------------------------DIALOG
BeginDialog homelessness_verified_dialog, 0, 0, 246, 105, "Homelessness Verified"
  EditBox 55, 5, 65, 15, MAXIS_case_number
  EditBox 190, 5, 50, 15, when_contact_was_made
  EditBox 55, 25, 65, 15, name_verf
  EditBox 190, 25, 50, 15, phone_number
  EditBox 55, 45, 185, 15, address_verf
  EditBox 55, 65, 185, 15, Comments_notes
  ButtonGroup ButtonPressed
    OkButton 135, 85, 50, 15
    CancelButton 190, 85, 50, 15
  Text 135, 30, 50, 10, "Phone number: "
  Text 165, 10, 20, 10, "Date:"
  Text 5, 10, 50, 10, "Case number: "
  Text 20, 50, 30, 10, "Address:"
  Text 15, 70, 40, 10, "Comments: "
  Text 25, 30, 25, 10, "Name:"
EndDialog

'--------------------------------------------------------------------------------------------------SCRIPT

EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'updates the "when contact was made" variable to show the current date & time
when_contact_was_made = date & ", " & time
DO
	Do
		Dialog homelessness_verified_dialog
		cancel_confirmation
		If (isnumeric(MAXIS_case_number) = False and len(MAXIS_case_number) <> 8) then MsgBox "You must enter either a valid MAXIS case number."
	Loop until (isnumeric(MAXIS_case_number) = True) or (isnumeric(MAXIS_case_number) = False and len(MAXIS_case_number) = 8)
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'----------------------------------------------------------------------------------------------------CASENOTE
start_a_blank_case_note
CALL write_variable_in_CASE_NOTE("### Homelessness Verified ###")
CALL write_bullet_and_variable_in_CASE_NOTE("Date:", when_contact_was_made)
CALL write_variable_in_CASE_NOTE("Name: " & name_verf)
CALL write_variable_in_CASE_NOTE("Phone number: " & phone_number)
CALL write_variable_in_CASE_NOTE("Address: " & address_verf)
CALL write_variable_in_CASE_NOTE("Comments: " & comments_notes)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team")


script_end_procedure("")