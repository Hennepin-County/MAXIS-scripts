'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - SUBMIT CASE FOR SNAP REVIEW.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 30           'manual run time in seconds
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

'Checks for county info from global variables, or asks if it is not already defined.
get_county_code

'THE DIALOG----------------------------------------------------------------------------------------------------
BeginDialog submitting_case_HENNEPIN_dialog, 0, 0, 271, 155, "Submitting Case for SNAP case review dialog"
  EditBox 70, 5, 60, 15, MAXIS_case_number					
  CheckBox 150, 10, 75, 10, "SNAP is expedited.", expedited_SNAP_check
  EditBox 70, 25, 195, 15, optional_info
  EditBox 70, 45, 85, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 160, 45, 50, 15
    CancelButton 215, 45, 50, 15
  Text 10, 10, 45, 10, "Case number: "
  Text 10, 50, 60, 10, "Worker signature: "
  GroupBox 5, 70, 260, 80, "Case submission reminders:"
  Text 10, 85, 245, 15, "* All cases with issuance amounts of $175 for Adults/ADS and $310 for FAD should be submitted for review."
  Text 10, 110, 245, 15, "* In CARL: click 'new search' before submitting a case, and select the first reviewer name."
  Text 10, 30, 45, 10, "Optional info:"
  Text 10, 135, 245, 15, "* In CARL: if SNAP is expedited, the put an 'E' before the case number."
EndDialog

BeginDialog submitting_case_dialog, 0, 0, 271, 70, "Submitting Case for SNAP case review dialog"
  EditBox 70, 5, 60, 15, MAXIS_case_number					
  CheckBox 150, 10, 75, 10, "SNAP is expedited.", expedited_SNAP_check
  EditBox 70, 25, 195, 15, optional_info
  EditBox 70, 45, 85, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 160, 45, 50, 15
    CancelButton 215, 45, 50, 15
  Text 10, 10, 45, 10, "Case number: "
  Text 10, 50, 60, 10, "Worker signature: "
  Text 10, 30, 45, 10, "Optional info:"
EndDialog


'The script----------------------------------------------------------------------------------------------------
'connecting to MAXIS & grabbing the case number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

Do 
	err_msg = ""
	If worker_county_code = "x127" THEN 
		Dialog submitting_case_HENNEPIN_dialog
	ELSE 
		Dialog submitting_case_dialog
	END IF
	If ButtonPressed = 0 then StopScript
	If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
	If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
LOOP until err_msg = ""

'checking for an active MAXIS session
Call check_for_MAXIS(False)

'The CASE NOTE----------------------------------------------------------------------------------------------------
Call start_a_blank_CASE_NOTE
If expedited_SNAP_check= 1 THEN 
	Call write_variable_in_CASE_NOTE("~~~~Case submitted for SNAP 2nd Review: EXPEDITED~~~~")
ELSE	
	Call write_variable_in_CASE_NOTE("~~~~Case submitted for SNAP 2nd Review~~~~")
END IF 
call write_bullet_and_variable_in_CASE_NOTE("Optional information", optional_info)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("")
