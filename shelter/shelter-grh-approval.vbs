'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - SHELTER-GRH APPROVAL.vbs"
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
'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 121, 105, "GRH Approval"
  EditBox 60, 5, 55, 15, MAXIS_case_number
  DropListBox 60, 25, 55, 15, "Select one..."+chr(9)+"FMF"+chr(9)+"PSP"+chr(9)+"SA-HL"+chr(9)+"St. Anne's"+chr(9)+"The Drake", shelter_droplist
  CheckBox 10, 65, 85, 10, "GRH notice cancelled", GRH_notice_checkbox
  EditBox 60, 45, 55, 15, ea_available_date
  ButtonGroup ButtonPressed
    OkButton 10, 85, 50, 15
    CancelButton 65, 85, 50, 15
  Text 10, 10, 50, 10, "Case Number:"
  Text 10, 50, 45, 10, "EA avail date:"
  Text 10, 30, 50, 10, "Shelter Name:"
EndDialog

'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		If shelter_droplist = "Select one..." then err_msg = err_msg & vbNewLine & "* Enter shelter name."
		If ea_available_date = "" then err_msg = err_msg & vbNewLine & "* Enter EA Available date"
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & "(enter NA in all fields that do not apply)" & vbNewLine & err_msg & vbNewLine
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
Call write_variable_in_CASE_NOTE("### GRH Approved ###")
Call write_bullet_and_variable_in_CASE_NOTE("Client placed in", shelter_droplist)
Call write_bullet_and_variable_in_CASE_NOTE("EA available date", ea_available_date)
If shelter_policy_checkbox = 1 then call write_variable_in_CASE_NOTE("* GRH Notice Cancelled.")
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE ("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team")

script_end_procedure("")
