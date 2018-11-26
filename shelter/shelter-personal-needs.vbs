'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - SHELTER-PERSONAL NEEDS.vbs"
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

'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog personal_needs_dialog, 0, 0, 171, 85, "Select a personal needs status"
  EditBox 85, 10, 55, 15, MAXIS_case_number
  DropListBox 85, 35, 70, 12, "Select one..."+chr(9)+"Approved client"+chr(9)+"Client ineligible", pers_needs_status
  ButtonGroup ButtonPressed
    OkButton 60, 60, 50, 15
    CancelButton 115, 60, 50, 15
  Text 5, 40, 75, 10, "Personal needs status:"
  Text 35, 15, 45, 10, "Case number:"
EndDialog

BeginDialog pers_needs_recd_dialog, 0, 0, 296, 70, "Personal needs received"
  EditBox 70, 20, 55, 15, amt_issued
  EditBox 225, 20, 55, 15, date_issued
  EditBox 55, 45, 115, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 175, 45, 50, 15
    CancelButton 230, 45, 50, 15
  Text 175, 25, 45, 10, "Date issued:"
  Text 10, 50, 40, 10, "Other notes: "
  Text 20, 25, 45, 10, "Amt issued:"
  GroupBox 5, 5, 280, 35, "Client received 10% + $20 personal needs:"
EndDialog

BeginDialog pers_needs_inelig_dialog, 0, 0, 296, 75, "Personal needs ineligible"
  EditBox 90, 25, 190, 15, inelig_reason
  ButtonGroup ButtonPressed
    OkButton 175, 55, 50, 15
    CancelButton 230, 55, 50, 15
  Text 10, 30, 80, 10, "Reason for ineligibility:"
  GroupBox 5, 10, 285, 35, "Client is ineligible for personal needs:"
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

date_issued = date & ""

'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog personal_needs_dialog
        cancel_confirmation
		IF len(case_number) > 8 or IsNumeric(case_number) = False THEN err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		IF pers_needs_status = "Select one..." then err_msg = err_msg & vbNewLine & "* Select a personal needs status."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
 Call check_for_password(are_we_passworded_out)
LOOP UNTIL check_for_password(are_we_passworded_out) = False

If pers_needs_status = "Approved client" then 
	case_note_header = "received"
	DO
		DO
			err_msg = ""
			Dialog pers_needs_recd_dialog
			cancel_confirmation
			IF IsNumeric(amt_issued) = False THEN err_msg = err_msg & vbNewLine & "* Enter the amount issued."
			IF IsDate(date_issued) = False then err_msg = err_msg & vbNewLine & "* Enter the date issued."
			If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP UNTIL err_msg = ""
 	Call check_for_password(are_we_passworded_out)
	LOOP UNTIL check_for_password(are_we_passworded_out) = False
END IF 

If pers_needs_status = "Client ineligible" then 
	case_note_header = "ineligible"
	DO
		DO
			err_msg = ""
			Dialog pers_needs_inelig_dialog
			cancel_confirmation
			IF len(case_number) > 8 or IsNumeric(case_number) = False THEN err_msg = err_msg & vbNewLine & "* Enter a valid case number."
			IF inelig_reason = "" then err_msg = err_msg & vbNewLine & "* Enter the personal needs ineligibility reason."
			If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
		LOOP UNTIL err_msg = ""
 		Call check_for_password(are_we_passworded_out)
	LOOP UNTIL check_for_password(are_we_passworded_out) = False
END IF 

back_to_SELF
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen CM_mo, 20, 43	'entering current footer month/year
EMWriteScreen CM_yr, 20, 46

amt_issued = "$" & amt_issued

'The case note---------------------------------------------------------------------------------------
start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
Call write_variable_in_CASE_NOTE("### Personal needs " & case_note_header & " ###")
If case_note_header = "received" then 
	Call write_variable_in_CASE_NOTE("* Client received 10% + $20 in personal needs funds. ")
	Call write_bullet_and_variable_in_CASE_NOTE("Amt issued", amt_issued)
	Call write_bullet_and_variable_in_CASE_NOTE("Date issued", date_issued)
ELSEIF case_note_header = "ineligible" then 
	Call write_bullet_and_variable_in_CASE_NOTE("Reason for ineligibility", inelig_reason)
END IF
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE ("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team")

script_end_procedure ("")