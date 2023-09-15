'GATHERING STATS===========================================================================================
name_of_script = "ACCT - Accounting Refund.vbs"
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
call changelog_update("06/20/2017", "Initial version.", "MiKayla Handley")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'--------------------------------------------------------------------------------------------------THE SCRIPT

EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 246, 135, "Shelter Refund"
  EditBox 65, 5, 65, 15, maxis_case_number
  EditBox 175, 5, 65, 15, when_contact_was_made
  EditBox 65, 30, 65, 15, check_number
  EditBox 175, 30, 65, 15, Check_amount
  DropListBox 80, 55, 160, 15, "Select One..."+chr(9)+"Mailed Out"+chr(9)+"Client Pickup  (North Service Desk)", Check_pickup_dropbox
  EditBox 80, 75, 160, 15, other_notes
  EditBox 80, 95, 160, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 135, 115, 50, 15
  Text 15, 35, 50, 10, "Check Number:"
  Text 15, 10, 50, 10, "Case Number:"
  Text 155, 10, 20, 10, "Date:"
  Text 35, 80, 40, 10, "Other notes:"
  Text 15, 60, 65, 10, "Method of Delivery:"
  Text 135, 35, 40, 10, "Check Amt:"
  ButtonGroup ButtonPressed
    CancelButton 190, 115, 50, 15
  Text 20, 100, 60, 10, "Worker Signature:"
EndDialog

'updates the "when contact was made" variable to show the current date & time
when_contact_was_made = date & ""
DO
	Do
        err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
		If (isnumeric(MAXIS_case_number) = False and len(MAXIS_case_number) <> 8) then err_msg = err_msg & vbnewline & "* You must enter either a valid MAXIS case number."
        If isDate(when_contact_was_made) = False then err_msg = err_msg & vbnewline & "* Enter a valid check date."
        If isnumeric(check_number) = False then err_msg = err_msg & vbnewline & "* Enter a valid numeric check number."
        If isnumeric(check_amount) = False then err_msg = err_msg & vbnewline & "* Enter a valid numeric check amount."
        If Check_pickup_dropbox = "Select One..." then err_msg = err_msg & vbnewline & "* Select a check pick up option."
        If worker_signature = "" then err_msg = err_msg & vbnewline & "* Enter your signature."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
	Loop until err_msg = ""
    call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'----------------------------------------------------------------------------------------------------CASENOTE
start_a_blank_case_note
CALL write_variable_in_CASE_NOTE("### Accounting Shelter Refund ###")
CALL write_variable_in_CASE_NOTE("Shelter Refund Issued on: " & when_contact_was_made)
CALL write_variable_in_CASE_NOTE("Check number: " & check_number &  ". Amount: $" & check_amount)
CALL write_variable_in_CASE_NOTE("Method of Delivery: " & Check_pickup_dropbox)
Call write_variable_in_CASE_NOTE("Other notes: " & other_notes)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE("Requested by, " & worker_signature)

script_end_procedure("")
