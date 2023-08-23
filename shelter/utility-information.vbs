'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - SHELTER-UTILITY INFO.vbs"
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
BeginDialog Dialog1, 0, 0, 331, 200, "Utilities information"
  EditBox 55, 5, 60, 15, MAXIS_case_number
  DropListBox 10, 60, 115, 15, "Select one..."+chr(9)+"CenterPoint Energy (VN #44)"+chr(9)+"Xcel Energy (VN #59499)"+chr(9)+"Water Company, MPLS (VN #394)"+chr(9)+"Other", vendor_type_1
  EditBox 140, 60, 60, 15, acct_number_1
  EditBox 210, 60, 45, 15, balance_1
  EditBox 265, 60, 50, 15, date_1
  DropListBox 10, 90, 115, 15, "Select one..."+chr(9)+"CenterPoint Energy (VN #44)"+chr(9)+"Xcel Energy (VN #59499)"+chr(9)+"Water Company, MPLS (VN #394)"+chr(9)+"Other", vendor_type_2
  EditBox 140, 90, 60, 15, acct_number_2
  EditBox 210, 90, 45, 15, balance_2
  EditBox 265, 90, 50, 15, date_2
  DropListBox 10, 120, 115, 15, "Select one..."+chr(9)+"CenterPoint Energy (VN #44)"+chr(9)+"Xcel Energy (VN #59499)"+chr(9)+"Water Company, MPLS (VN #394)"+chr(9)+"Other", vendor_type_3
  EditBox 140, 120, 60, 15, acct_number_3
  EditBox 210, 120, 45, 15, balance_3
  EditBox 265, 120, 50, 15, date_3
  DropListBox 10, 150, 115, 15, "Select one..."+chr(9)+"CenterPoint Energy (VN #44)"+chr(9)+"Xcel Energy (VN #59499)"+chr(9)+"Water Company, MPLS (VN #394)"+chr(9)+"Other", vendor_type_4
  EditBox 140, 150, 60, 15, acct_number_4
  EditBox 210, 150, 45, 15, balance_4
  EditBox 265, 150, 50, 15, date_4
  EditBox 165, 5, 150, 15, other_information
  EditBox 75, 180, 140, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 220, 180, 50, 15
    CancelButton 275, 180, 50, 15
  Text 150, 40, 35, 10, "Account #"
  Text 5, 10, 50, 10, "Case Number"
  Text 30, 40, 55, 10, "Utility company"
  Text 220, 40, 30, 10, "Balance"
  Text 125, 10, 40, 10, "Comments:"
  Text 265, 40, 45, 10, "Balance date"
  GroupBox 5, 25, 320, 150, "Complete for each of the client(s) utility:"
  Text 5, 185, 60, 10, "Worker Signature:"
EndDialog

DO
	DO
		err_msg = ""
		Dialog Dialog1
        cancel_confirmation
		IF len(MAXIS_case_number) > 8 or IsNumeric(MAXIS_case_number) = False THEN err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		IF (vendor_type_1 = "Select one..." AND vendor_type_2 = "Select one..." AND vendor_type_3 = "Select one..." AND vendor_type_4 = "Select one...") THEN err_msg = err_msg & vbCr & "*At least one vendor is needed."
		IF (vendor_type_1 <> "Select one..." AND (vendor_number_1 = "" AND acct_number_1 = "" AND balance_1 = "" AND date_1 = "")) THEN err_msg = err_msg & vbCr & "*All vendor information must be completed for the 1st vendor selected."
		IF (vendor_type_2 <> "Select one..." AND (vendor_number_2 = "" AND acct_number_2 = "" AND balance_2 = "" AND date_2 = "")) THEN err_msg = err_msg & vbCr & "*All vendor information must be completed for the 2nd vendor selected."
		IF (vendor_type_3 <> "Select one..." AND (vendor_number_3 = "" AND acct_number_3 = "" AND balance_3 = "" AND date_3 = "")) THEN err_msg = err_msg & vbCr & "*All vendor information must be completed for the 3rd vendor selected."
		IF (vendor_type_4 <> "Select one..." AND (vendor_number_4 = "" AND acct_number_4 = "" AND balance_4 = "" AND date_4 = "")) THEN err_msg = err_msg & vbCr & "*All vendor information must be completed for the 4th vendor selected."
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

back_to_SELF
EMWriteScreen "________", 18, 43
EMWriteScreen MAXIS_case_number, 18, 43
EMWriteScreen CM_mo, 20, 43	'entering current footer month/year
EMWriteScreen CM_yr, 20, 46

'The case note---------------------------------------------------------------------------------------
start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
Call write_variable_in_CASE_NOTE("### Utility information ###")
Call write_variable_in_CASE_NOTE ("Client has a bill with the following companies:")
Call write_variable_in_CASE_NOTE ("------------------------------------------------")
IF vendor_type_1 <> "Select one..." then
    Call write_variable_in_CASE_NOTE(vendor_type_1)
	Call write_variable_in_CASE_NOTE("   * Acct #: " & acct_number_1)
    Call write_variable_in_CASE_NOTE("   * Balance: $" & balance_1 & " as of " & date_1)
    Call write_variable_in_CASE_NOTE ("-")
END IF
IF vendor_type_2 <> "Select one..." then
	Call write_variable_in_CASE_NOTE(vendor_type_2)
	Call write_variable_in_CASE_NOTE("   * Acct #: " & acct_number_2)
	Call write_variable_in_CASE_NOTE("   * Balance: $" & balance_2 & " as of " & date_2)
	Call write_variable_in_CASE_NOTE ("-")
END IF

IF vendor_type_3 <> "Select one..." then
	Call write_variable_in_CASE_NOTE(vendor_type_3)
	Call write_variable_in_CASE_NOTE("   * Acct #: " &  acct_number_3)
	Call write_variable_in_CASE_NOTE("   * Balance: $" & balance_3 & " as of " & date_3)
	Call write_variable_in_CASE_NOTE ("-")
END IF

IF vendor_type_4 <> "Select one..." then
	Call write_variable_in_CASE_NOTE(vendor_type_4)
	Call write_variable_in_CASE_NOTE("   * Acct #: " & acct_number_4)
	Call write_variable_in_CASE_NOTE("   * Balance: $" & balance_4 & " as of " & date_4)
	Call write_variable_in_CASE_NOTE ("-")
END IF

Call write_bullet_and_variable_in_CASE_NOTE("Other Information", other_information)
Call write_variable_in_CASE_NOTE ("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team")

script_end_procedure("")
