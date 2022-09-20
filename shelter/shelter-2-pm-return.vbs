'GATHERING STATS===========================================================================================
name_of_script = "NOTES - SHELTER-2 PM RETURN.vbs"
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
CALL changelog_update("09/20/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
call changelog_update("06/26/2017", "Initial version.", "MiKayla Handley")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'--------------------------------------------------------------------------------------------------THE SCRIPT

EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 186, 105, " 2pm Return "
  EditBox 70, 5, 50, 15, maxis_case_number
  EditBox 70, 25, 110, 15, Comments_notes
  EditBox 70, 45, 110, 15, items_still_needed
  EditBox 70, 65, 110, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 70, 85, 50, 15
    CancelButton 130, 85, 50, 15
  Text 5, 10, 50, 10, "Case Number:"
  Text 5, 50, 65, 10, "Items Still Needed:"
  Text 5, 30, 40, 10, "Comments:"
  Text 5, 75, 60, 10, "Worker signature:"
EndDialog
DO
	Do
        err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
		Call validate_MAXIS_case_number(err_msg, "*")
        If worker_signature = "" then err_msg = err_msg & vbnewline & "* Enter your signature."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = ""
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
CALL write_variable_in_CASE_NOTE("### 2pm Return ###")
CALL write_variable_in_CASE_NOTE("Client returning at 2pm and requesting shelter. ")
CALL write_variable_in_CASE_NOTE("Comments: " & Comments_notes)
CALL write_variable_in_CASE_NOTE("Items Still Needed:" & items_still_needed)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE("Requested by, " & worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team")

script_end_procedure("")
