'GATHERING STATS===========================================================================================
name_of_script = "SHELTER-CASH CUT-OFF.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 180
STATS_denominatinon = "C"
'END OF STATS BLOCK===========================================================================================

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
call changelog_update("06/20/2017", "Initial version.", "MiKayla Handley")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'-------------------------------------------------------------------------------------------------DIALOG
BeginDialog Cash_Cut_Off_dlg, 0, 0, 186, 95, " Cash Cut-Off "
  EditBox 60, 5, 65, 15, maxis_case_number
  DropListBox 130, 30, 50, 15, "Select One..."+chr(9)+"YES"+chr(9)+"NO", VNDS_dropbox
  EditBox 45, 55, 135, 15, Comments_notes
  ButtonGroup ButtonPressed
    OkButton 75, 75, 50, 15
    CancelButton 130, 75, 50, 15
  Text 5, 60, 40, 10, "Comments:"
  Text 5, 10, 50, 10, "Case Number:"
  Text 5, 35, 120, 10, "All cash vendored to shelter account:"
EndDialog

'--------------------------------------------------------------------------------------------------THE SCRIPT

EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'updates the "when contact was made" variable to show the current date & time
when_contact_was_made = date 
DO
	Do
		Dialog Cash_Cut_Off_dlg
		cancel_confirmation
		If (isnumeric(MAXIS_case_number) = False and len(MAXIS_case_number) <> 8) then MsgBox "You must enter either a valid MAXIS case number."
	Loop until (isnumeric(MAXIS_case_number) = True) or (isnumeric(MAXIS_case_number) = False and len(MAXIS_case_number) = 8)
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false
back_to_SELF
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen CM_mo, 20, 43	 '------------------------------------------------------------------------------''entering current footer month/year
EMWriteScreen CM_yr, 20, 46
date_header = CM_mo & "/" & CM_yr

'----------------------------------------------------------------------------------------------------CASENOTE
start_a_blank_case_note
CALL write_variable_in_CASE_NOTE("### CASH CUT OFF MONY/VNDA checked for " & date_header & " ###")
CALL write_variable_in_CASE_NOTE("All cash vendored to shelter account? " & VNDS_dropbox)
CALL write_variable_in_CASE_NOTE("Comments: " & Comments_notes)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team")

script_end_procedure("")