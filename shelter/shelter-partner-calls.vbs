'GATHERING STATS===========================================================================================
name_of_script = "NOTES - SHELTER-PARTNER CALLS.vbs"
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
call changelog_update("06/19/2017", "Initial version.", "MiKayla Handley")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'-------------------------------------------------------------------------------------------------DIALOG
BeginDialog partner_calls_dialog, 0, 0, 306, 235, "Partner Calls"
  EditBox 230, 165, 65, 15, maxis_case_number
  EditBox 230, 185, 65, 15, when_contact_was_made
  EditBox 80, 20, 65, 15, ESP_Name
  EditBox 80, 40, 65, 15, ESP_Phone_number
  EditBox 80, 60, 65, 15, ESP_FSS_comments
  EditBox 80, 95, 65, 15, CP_name
  EditBox 80, 115, 65, 15, CP_phone_number
  EditBox 80, 135, 65, 15, CP_Comments
  EditBox 230, 20, 65, 15, PO_name
  EditBox 230, 40, 65, 15, PO_phone_number
  EditBox 230, 60, 65, 15, PO_comments
  EditBox 230, 95, 65, 15, RRH_name
  EditBox 230, 115, 65, 15, RRH_phone_number
  EditBox 230, 135, 65, 15, RRH_comment
  EditBox 80, 170, 65, 15, OTHER_name
  EditBox 80, 190, 65, 15, Other_phone_number
  EditBox 80, 210, 65, 15, Other_comments
  ButtonGroup ButtonPressed
    OkButton 190, 215, 50, 15
    CancelButton 245, 215, 50, 15
  Text 15, 175, 65, 10, "Organization/Name: "
  GroupBox 10, 160, 140, 70, "OTHER:"
  Text 15, 215, 45, 10, "Comments:"
  GroupBox 10, 10, 140, 70, "ESP/FSS:"
  GroupBox 10, 85, 140, 70, "CP:"
  Text 15, 100, 65, 10, "Organization/Name: "
  Text 165, 45, 50, 10, "Phone number: "
  Text 165, 25, 65, 10, "Organization/Name: "
  GroupBox 160, 10, 140, 70, "PO:"
  Text 165, 65, 45, 10, "Comments:"
  Text 15, 25, 65, 10, "Organization/Name: "
  Text 15, 45, 50, 10, "Phone number: "
  Text 15, 140, 45, 10, "Comments:"
  Text 165, 120, 50, 10, "Phone number: "
  Text 165, 100, 65, 10, "Organization/Name: "
  GroupBox 160, 85, 140, 70, "RRH:"
  Text 165, 140, 45, 10, "Comments:"
  Text 15, 65, 45, 10, "Comments:"
  Text 175, 170, 50, 10, "Case Number:"
  Text 15, 120, 50, 10, "Phone number: "
  Text 205, 190, 20, 10, "Date:"
  Text 15, 195, 50, 10, "Phone number: "
EndDialog

'--------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'updates the "when contact was made" variable to show the current date & time
when_contact_was_made = date & ", " & time
DO
	Do
		Dialog partner_calls_dialog
		cancel_confirmation
		If (isnumeric(MAXIS_case_number) = False and len(MAXIS_case_number) <> 8) then MsgBox "You must enter either a valid MAXIS case number."
	Loop until (isnumeric(MAXIS_case_number) = True) or (isnumeric(MAXIS_case_number) = False and len(MAXIS_case_number) = 8)
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'----------------------------------------------------------------------------------------------------CASENOTE
start_a_blank_case_note
CALL write_variable_in_CASE_NOTE("### Partner Calls ###")
CALL write_bullet_and_variable_in_CASE_NOTE("Contact was made", when_contact_was_made)
IF ESP_name <> "" THEN 
    CALL write_variable_in_CASE_NOTE("* ESP/FSS Organization/Name: " & ESP_name)
    CALL write_variable_in_CASE_NOTE("* ESP/FSS Phone number: " & ESP_Phone_number)
    CALL write_variable_in_CASE_NOTE("* Comments: " & ESP_FSS_comments)
	Call write_variable_in_CASE_NOTE("---")
END IF
IF CP_name <> "" THEN 
    CALL write_variable_in_CASE_NOTE("* CP Organization/Name: " & CP_name)
    CALL write_variable_in_CASE_NOTE("* CP Phone number: " & CP_Phone_number)
    CALL write_variable_in_CASE_NOTE("* Comments: " & CP_comments)
	Call write_variable_in_CASE_NOTE("---")
END IF
IF PO_name <> "" THEN 
    CALL write_variable_in_CASE_NOTE("* PO Organization/Name: " & PO_name)
    CALL write_variable_in_CASE_NOTE("* PO Phone number: " & PO_Phone_number)
    CALL write_variable_in_CASE_NOTE("* Comments: " & PO_comments)
	Call write_variable_in_CASE_NOTE("---")
END IF
IF RRH_name <> "" THEN 
    CALL write_variable_in_CASE_NOTE("* Rapid Re-Housing Organization/Name: " & RRH_name)
    CALL write_variable_in_CASE_NOTE("* RRH Phone number: " & RRH_Phone_number)
    CALL write_variable_in_CASE_NOTE("* Comments: " & RRH_comments)
	Call write_variable_in_CASE_NOTE("---")
END IF
IF OTHER_name <> "" THEN 
    CALL write_variable_in_CASE_NOTE("* Organization/Name: " & OTHER_name)
    CALL write_variable_in_CASE_NOTE("* Phone number: " & Other_Phone_number)
    CALL write_variable_in_CASE_NOTE("* Comments: " & Other_comments)
Call write_variable_in_CASE_NOTE("---")
END IF
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team")


script_end_procedure("")