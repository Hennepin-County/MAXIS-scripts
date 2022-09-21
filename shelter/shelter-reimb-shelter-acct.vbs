'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - SHELTER-REIM SHELTER ACCT.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 0         	'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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
CALL changelog_update("09/21/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
call changelog_update("06/19/2017", "Initial version.", "MiKayla Handley")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 271, 120, "### Reimbursement from shelter account ###"
  EditBox 55, 5, 45, 15, MAXIS_case_number
  EditBox 225, 5, 40, 15, amt_to_client
  EditBox 225, 25, 40, 15, amt_to_LL
  EditBox 70, 45, 195, 15, refund_reason
  CheckBox 5, 60, 240, 10, "Informed client funds are being released for basic needs for the month,", client_checkbox
  EditBox 70, 80, 195, 15, other_notes
  EditBox 70, 100, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 175, 100, 45, 15
    CancelButton 220, 100, 45, 15
  Text 145, 10, 80, 10, "Amt released to client $:"
  Text 5, 85, 40, 10, "Comments: "
  Text 5, 10, 45, 10, "Case number:"
  Text 15, 70, 175, 10, "and shelter will not be available for the entire month."
  Text 100, 30, 125, 10, "Amt vendored to LL for rent/deposit $:"
  Text 5, 50, 50, 10, "Refund reason:"
  Text 5, 105, 60, 10, "Worker signature:"
EndDialog
'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
		If amt_to_client = "" then err_msg = err_msg & vbNewLine & "* Please enter the amount released to the client."
		If amt_to_LL = "" then err_msg = err_msg & vbNewLine & "* Please enter the amount vendored to the landlord."
		If refund_reason = "" then err_msg = err_msg & vbNewLine & "* Please enter the refund reason."
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'adding the case number
back_to_self
EMWriteScreen "________", 18, 43
EMWriteScreen MAXIS_case_number, 18, 43
EMWriteScreen CM_mo, 20, 43	'entering current footer month/year
EMWriteScreen CM_yr, 20, 46

'The case note'
start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("### HCEA-ACF shelter account ###")
Call write_bullet_and_variable_in_CASE_NOTE("Refund reason", refund_reason)
Call write_bullet_and_variable_in_CASE_NOTE("Amount released to client", amt_to_client)
Call write_bullet_and_variable_in_CASE_NOTE("Amount vendored to LL for rent/deposit", amt_to_LL)
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
If client_checkbox = 1 then Call write_variable_in_CASE_NOTE("* Informed client funds are being released for basic needs for the month, and shelter will not be available for the entire month." )
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team")

script_end_procedure("")
