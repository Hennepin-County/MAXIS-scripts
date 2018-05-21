'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - SHELTER-SINGLE CLIENT INTERVIEW.vbs"
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
BeginDialog single_client_interview_dialog, 0, 0, 301, 245, "Single client interview"
  EditBox 70, 10, 55, 15, MAXIS_case_number
  DropListBox 190, 10, 60, 15, "Select one..."+chr(9)+"PSP"+chr(9)+"SA-HL", shelter_type
  EditBox 70, 35, 30, 15, num_nights
  EditBox 190, 35, 95, 15, shelter_dates
  CheckBox 10, 60, 285, 10, "Explained shelter policies, self pay and client options to shleter such as bus tickets,", shelter_policy_checkbox
  CheckBox 10, 90, 115, 10, "ATR's and data sharing signed.", ATR_checkbox
  CheckBox 10, 110, 280, 10, "MOF given to client to have Dr. complete/return to HSR team to determine GA basis.", MOF_checkbox
  CheckBox 10, 130, 115, 10, "ATR's and data sharing signed.", Check4
  CheckBox 10, 150, 255, 10, "(18-21 YRS): Form given to client to take to Margo to determine school plan,", school_plan_checkbox
  CheckBox 10, 180, 185, 10, "Set TIKL for revoucher date.               Revoucher date:", set_TIKL
  EditBox 200, 175, 55, 15, revoucher_date
  EditBox 70, 200, 220, 15, GA_pending
  EditBox 70, 220, 105, 15, other_notes
  ButtonGroup ButtonPressed
    OkButton 185, 220, 50, 15
    CancelButton 240, 220, 50, 15
  Text 20, 70, 135, 10, " temporary housing, private shelters, etc. "
  Text 30, 40, 40, 10, "# of nights:"
  Text 15, 205, 50, 10, "GA pending for:"
  Text 110, 40, 80, 10, "Dates shelter issued for:"
  Text 30, 225, 40, 10, "Comments: "
  Text 20, 160, 175, 10, " and return by the revoucher date in order to get more voucher."
  Text 20, 15, 45, 10, "Case number:"
  Text 145, 15, 45, 10, "Shelter type:"
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to BlueZone, grabbing case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Running the initial dialog
DO
	DO
		err_msg = ""
		Dialog single_client_interview_dialog
        cancel_confirmation
		IF len(MAXIS_case_number) > 8 or IsNumeric(MAXIS_case_number) = False THEN err_msg = err_msg & vbNewLine & "* Enter a valid case number."
		IF shelter_type = "Select one..." then err_msg = err_msg & vbNewLine & "* Select a voucher type."
		If IsNumeric(num_nights) = False then err_msg = err_msg & vbNewLine & "* Enter the number of nights issued."
		If shelter_dates = "" then err_msg = err_msg & vbNewLine & "* Enter the dates of the shelter stay."
		If set_TIKL = 1 and isDate(revoucher_date) = False then err_msg = err_msg & vbNewLine & "* Please enter the revoucher date for the TIKL." 
		If GA_pending = "" then err_msg = err_msg & vbNewLine & "* Enter the reason GA is pending."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP UNTIL err_msg = ""
 Call check_for_password(are_we_passworded_out)
LOOP UNTIL check_for_password(are_we_passworded_out) = False

back_to_SELF
EMWriteScreen "________", 18, 43
EMWriteScreen case_number, 18, 43
EMWriteScreen CM_mo, 20, 43	'entering current footer month/year
EMWriteScreen CM_yr, 20, 46

'Creates a TIKL for the revoucher date
If set_TIKL = 1 then
	call navigate_to_MAXIS_screen("dail", "writ")
	call create_MAXIS_friendly_date(revoucher_date, 0, 5, 18) 
	Call write_variable_in_TIKL("Revoucher date. Please review case for requested verifications and/or redetermination of benefits.")
	transmit	
	PF3
End if

'The case note---------------------------------------------------------------------------------------
start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
Call write_variable_in_CASE_NOTE("### CLIENT APPROVED FOR SHELTER ###")
Call write_variable_in_CASE_NOTE("* Client has been housed at a " & shelter_type & " shelter for " & num_nights & " nights.")
Call write_bullet_and_variable_in_CASE_NOTE("Dates shelter has been issued for", shelter_dates)
If shelter_policy_checkbox = 1 then call write_variable_in_CASE_NOTE("* Explained shelter policies, self pay and client option to shelter such as bus tickets, temporary housing, private shelters, etc.")
If ATR_checkbox = 1 then call write_variable_in_CASE_NOTE("* ATR and data sharing signed.")
If MOF_checkbox = 1 then call write_variable_in_CASE_NOTE("* MOF given to client to have Dr. complete and return to HSR team to determine GA basis.")
If school_plan_checkbox = 1 then call write_variable_in_CASE_NOTE("* Form given to client to take to take to Margo to determine school plan, and return by the revoucher date in order to get more voucher.")
If set_TIKL = 1 then Call write_bullet_and_variable_in_CASE_NOTE("Set TIKL for revoucher date of", revoucher_date)
Call write_variable_in_CASE_NOTE ("---")
Call write_bullet_and_variable_in_CASE_NOTE("GA pending for", GA_pending)
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_variable_in_CASE_NOTE ("---")
Call write_variable_in_CASE_NOTE(worker_signature)
Call write_variable_in_CASE_NOTE("Hennepin County Shelter Team")

If set_TIKL = 1 then 
	script_end_procedure("A TIKL has been set for " & revoucher_date & " to recheck case.")
ELSE 
	script_end_procedure("")
END IF