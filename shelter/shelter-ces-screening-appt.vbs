'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - SHELTER-CES SCREENING APPT.vbs"
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
BeginDialog CES_screening_appt, 0, 0, 271, 100, "CES Screening Appointment Scheduled"
  EditBox 70, 10, 55, 15, MAXIS_case_number
  EditBox 210, 10, 55, 15, memb_name
  EditBox 70, 30, 55, 15, appt_date
  EditBox 210, 30, 55, 15, appt_time
  CheckBox 5, 50, 225, 10, "Informed client to bring 3 year rental/ADDR history, and meet with", informed_client
  EditBox 45, 75, 110, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 160, 75, 50, 15
    CancelButton 215, 75, 50, 15
  Text 25, 60, 190, 10, "the Shelter Team for initial interview after CES screening."
  Text 5, 80, 40, 10, "Comments:"
  Text 5, 35, 60, 10, "Appointment date:"
  Text 20, 15, 45, 10, "Case number:"
  Text 145, 35, 60, 10, "Appointment time:"
  Text 155, 15, 50, 10, "Member name:"
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog CES_screening_appt
        cancel_confirmation
		IF len(MAXIS_case_number) > 8 or IsNumeric(MAXIS_case_number) = False THEN err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		IF memb_name = "" then err_msg = err_msg & vbNewLine & "* Enter the referred member's name."
		If Isdate(appt_date) = False then err_msg = err_msg & vbNewLine & "* Enter the CES appointment date."
		If appt_time = "" then err_msg = err_msg & vbNewLine & "* Enter the CES appointment time."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
 Call check_for_password(are_we_passworded_out)
LOOP UNTIL check_for_password(are_we_passworded_out) = False
 
back_to_SELF

'The case note---------------------------------------------------------------------------------------
start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
Call write_variable_in_CASE_NOTE("--CES Screening Appt. scheduled for " & memb_name & " on " & appt_date & " at " & appt_time & "--")
If informed_client = 1 then Call write_variable_in_CASE_NOTE("* Informed client to bring 3 years of rental/ADDR history, and meet with the Shelter Team for an initial interview after the CES screening.")
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team")

script_end_procedure("")