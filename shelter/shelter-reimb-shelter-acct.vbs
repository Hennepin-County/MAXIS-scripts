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

'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog reimb_shel_acct, 0, 0, 296, 125, "### Reimbursement from shelter account ###"
  EditBox 65, 10, 55, 15, MAXIS_case_number
  EditBox 240, 10, 50, 15, amt_to_client
  EditBox 240, 30, 50, 15, amt_to_LL
  EditBox 65, 55, 225, 15, refund_reason
  CheckBox 10, 80, 280, 10, "Informed client funds are being released for basic needs for the month, and shelter", client_checkbox
  EditBox 50, 105, 125, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 185, 105, 50, 15
    CancelButton 240, 105, 50, 15
  Text 155, 15, 80, 10, "Amt released to client $:"
  Text 10, 110, 40, 10, "Comments: "
  Text 10, 15, 45, 10, "Case number:"
  Text 20, 90, 140, 10, "will not be available for the entire month."
  Text 105, 35, 125, 10, "Amt vendored to LL for rent/deposit $:"
  Text 10, 60, 50, 10, "Refund reason:"
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog reimb_shel_acct
		cancel_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		If amt_to_client = "" then err_msg = err_msg & vbNewLine & "* Enter the amount released to the client."
		If amt_to_LL = "" then err_msg = err_msg & vbNewLine & "* Enter the amount vendored to the landlord."
		If refund_reason = "" then err_msg = err_msg & vbNewLine & "* Enter the refund reason."
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