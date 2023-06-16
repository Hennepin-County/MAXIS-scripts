'GATHERING STATS===========================================================================================
name_of_script = "NOTES - ABAWD TRACKING RECORD.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 240
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
call changelog_update("04/19/2021", "Removed SNAP Banked Months information as it is no longer valid.", "Ilse Ferris, Hennepin County")
call changelog_update("07/17/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to MAXIS & case number
EMConnect ""
call maxis_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
member_number = "01"
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 141, 95, "Enter the case number & footer month/year"
  EditBox 75, 5, 45, 15, MAXIS_case_number
  EditBox 75, 25, 20, 15, MAXIS_footer_month
  EditBox 100, 25, 20, 15, MAXIS_footer_year
  EditBox 75, 45, 20, 15, member_number
  ButtonGroup ButtonPressed
	OkButton 15, 65, 50, 15
	CancelButton 70, 65, 50, 15
  Text 5, 30, 65, 10, "Footer month/year:"
  Text 20, 10, 55, 10, "Case Number: "
  Text 35, 50, 40, 10, "Member #:"
EndDialog
Do
	Do
	    err_msg = ""
  		Dialog Dialog1
  		cancel_confirmation
  		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
  		If IsNumeric(MAXIS_footer_month) = False or len(MAXIS_footer_month) > 2 or len(MAXIS_footer_month) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer month."
  		If IsNumeric(MAXIS_footer_year) = False or len(MAXIS_footer_year) > 2 or len(MAXIS_footer_year) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer year."
		If IsNumeric(member_number) = False or len(member_number) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer year."
  		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

MAXIS_background_check
MAXIS_footer_month_confirmation
Call navigate_to_MAXIS_screen("STAT", "WREG")
Do
	EMReadScreen WREG_panel, 4, 2, 48
	If WREG_panel <> "WREG" then Call navigate_to_MAXIS_screen("STAT", "WREG")
Loop until WREG_panel = "WREG"

CALL write_value_and_transmit(member_number, 20, 76)
CALL write_value_and_transmit("x", 13, 57)

EMReadScreen ATR_header, 65, 4, 11
EMReadScreen ATR_months, 53, 6, 12
EMReadScreen ATR_line_one, 52, 7, 12
EMReadScreen ATR_line_two, 52, 8, 12
EMReadScreen ATR_line_three, 52, 9, 12
EMReadScreen ATR_line_four, 52, 10, 12

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 356, 125, "ABAWD Tracking Record for MEMB  & member_number"
  EditBox 65, 20, 280, 15, ABAWD_months
  EditBox 65, 40, 280, 15, second_months
  EditBox 65, 65, 280, 15, deleted_months
  EditBox 65, 85, 280, 15, other_notes
  EditBox 65, 105, 175, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 245, 105, 50, 15
    CancelButton 295, 105, 50, 15
  Text 25, 45, 40, 10, "Second set:"
  Text 10, 25, 55, 10, "ABAWD months:"
  Text 20, 90, 45, 10, "Other Notes:"
  GroupBox 5, 5, 345, 55, "Please detail information about this resident's ABAWD Tracking Record below:"
  Text 5, 110, 60, 10, "Worker Signature:"
  Text 10, 70, 55, 10, "Deleted months:"
EndDialog

'the dialog
Do
	Do
		err_msg = ""
  		Dialog Dialog1
  		cancel_confirmation
        Call validate_MAXIS_case_number(err_msg, "*")
        Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")
		If trim(worker_signature) = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
  		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

PF3 'Exiting the ABAWD tracking record
'----------------------------------------------------------------------------------------------------The case note
start_a_blank_CASE_NOTE
call write_variable_in_CASE_NOTE("**Updated ABAWD tracking record for MEMB " & member_number & "**")
call write_variable_in_CASE_NOTE(ATR_header)
call write_variable_in_CASE_NOTE(ATR_months)
call write_variable_in_CASE_NOTE(ATR_line_one)
call write_variable_in_CASE_NOTE(ATR_line_two)
call write_variable_in_CASE_NOTE(ATR_line_three)
call write_variable_in_CASE_NOTE(ATR_line_four)
call write_bullet_and_variable_in_CASE_NOTE("ABAWD months", ABAWD_months)
call write_bullet_and_variable_in_CASE_NOTE("2nd set ABAWD months", Second_months)
call write_bullet_and_variable_in_CASE_NOTE("Deleted months", deleted_months)
call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(Worker_Signature)

script_end_procedure("")
