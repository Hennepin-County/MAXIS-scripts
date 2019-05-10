'Required for statistical purposes===============================================================================
name_of_script = "DAIL - MEDI CHECK.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 127         'manual run time in seconds
STATS_denomination = "C"       'C is for case
'END OF stats block==============================================================================================

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

'===========================================================================================================CHANGELOG BLOCK
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("05/01/2019", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'=======================================================================================================END CHANGELOG BLOCK
'THE MAIN DIALOG--------------------------------------------------------------------------------------------------
BeginDialog other_bene_dialog, 0, 0, 316, 215, "Other Maintenance Benefits" & maxis_case_number
  EditBox 70, 15, 15, 15, memb_number
  EditBox 185, 15, 50, 15, other_elig_date
  CheckBox 15, 30, 80, 10, "Railroad Retirement", railroad_retirement_checkbox
  CheckBox 15, 40, 90, 10, "Worker's Compensation", worker_compensation_checkbox
  CheckBox 15, 50, 95, 10, "Unemployment Insurance", unemployment_insurance_checkbox
  CheckBox 15, 60, 130, 10, "Other (please specify in other notes)", other_checkbox
  CheckBox 130, 30, 180, 10, "Retirement, Survivors, and Disability Income (RSDI)", RSDI_checkbox
  CheckBox 130, 40, 125, 10, "Supplemental Security Income (SSI)", SSI_checkbox
  CheckBox 130, 50, 120, 10, "Veterans' Disability Benefits (VA)", VA_checkbox
  CheckBox 15, 85, 65, 10, "Medicare Buy-In", medi_checkbox
  EditBox 70, 95, 50, 15, ELIG_date
  EditBox 290, 95, 15, 15, ELIG_year
  CheckBox 10, 115, 235, 10, "Sent Notice to Apply for Other Maintenance Benefits DHS-2116-ENG", ECF_sent_checkbox
  CheckBox 10, 125, 285, 10, "Client will need to fill out an Interim Assistance Agreement DHS-1795 or DHS-1795A", IAA_needed
  EditBox 65, 140, 240, 15, action_taken
  EditBox 65, 160, 240, 15, other_notes
  EditBox 65, 180, 150, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 220, 180, 40, 15
    CancelButton 265, 180, 40, 15
  GroupBox 5, 5, 305, 70, "Client may be able to get  benefits from the programs checked below:"
  GroupBox 5, 75, 305, 40, "Medicare Buy-In only **   "
  Text 15, 100, 50, 10, "Eligibility Date:"
  Text 140, 100, 150, 10, "If ineligible the year that the client will be elig:"
  Text 15, 145, 50, 10, "Actions Taken:"
  Text 20, 165, 45, 10, "Other Notes:"
  Text 5, 185, 60, 10, "Worker signature:"
  Text 290, 85, 15, 10, "(YY)"
  Text 130, 20, 50, 10, "Eligibility Date:"
  Text 15, 20, 50, 10, "Memb Number:"
EndDialog



EMConnect ""

EMWriteScreen "N", 6, 3         'Goes to Case Note - maintains tie with DAIL
TRANSMIT
'Starts a blank case note
PF9
EMReadScreen case_note_mode_check, 7, 20, 3
If case_note_mode_check <> "Mode: A" then script_end_procedure("You are not in a case note on edit mode. You might be in inquiry. Try the script again in production.")

Do
    Do
        err_msg = ""
		Dialog medi_dialog
		cancel_confirmation
		IF (isnumeric(memb_number) = False and len(memb_number) > 2) then err_msg = err_msg & vbcr & "* Enter a valid member number."
		IF medi_checkbox = CHECKED THEN If isdate(ELIG_date) = False then err_msg = err_msg & vbnewline & "* Enter a valid date of eligibility."
        If (isnumeric(MAXIS_case_number) = False and len(MAXIS_case_number) <> 8) then err_msg = err_msg & vbcr & "* Enter a valid case number."
		If trim(worker_signature) = "" then err_msg = err_msg & vbcr & "* Please ensure your case note is signed."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

PF9
'CALL write_variable_in_CASE_NOTE("=== PEPR - MESSAGE PROCESSED ===")
'CALL write_variable_in_case_note("* " & full_message)
'CALL write_variable_in_case_note(first_line)
'CALL write_variable_in_case_note(second_line)
'CALL write_variable_in_case_note(third_line)
'CALL write_variable_in_case_note(fourth_line)
'CALL write_variable_in_case_note(fifth_line)
'CALL write_variable_in_case_note("---")
IF medi_checkbox = CHECKED THEN
	due_date = dateadd("d", 30, date)
	Call write_variable_in_case_note("** Medicare Buy-in Referral mailed for M" & memb_number & " **")
	Call write_variable_in_case_note("Client is eligible for the Medicare buy-in as of " & ELIG_date & ". Proof due by " & due_date & " to apply.")
	Call write_variable_in_case_note("Mailed DHS-3439-ENG MHCP Medicare Buy-In Referral Letter - TIKL set to follow up.")
ELSE
	Call write_variable_in_case_note("** Medicare Referral for M" & memb_number & " **")
	Call write_variable_in_case_note("Client is not eligible for the Medicare buy-in. Enrollment is not until January " & ELIG_year & ", unable to apply until the enrollment time.")
	Call write_variable_in_case_note("TIKL set to mail the Medicare Referral for November " & ELIG_year & ".")
END IF
IF ECF_sent_checkbox = CHECKED THEN CALL write_variable_in_case_note("* ECF reviewed and appropriate action taken")
CALL write_bullet_and_variable_in_case_note("Other notes", other_notes)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)
PF3


'TIKLING
IF medi_checkbox = CHECKED and ELIG_date <> "" THEN
	CALL navigate_to_MAXIS_screen("DAIL","WRIT")
	call create_MAXIS_friendly_date(Due_date, 10, 5, 18)
	CALL write_variable_in_TIKL("Referral made for medicare, please check on proof of application filed. Due " & due_date & ".")
	PF3
END IF
IF ELIG_year <> "" THEN
	CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
	EMWriteScreen "11", 5, 18
	EMWriteScreen "01", 5, 21
	EMWriteScreen ELIG_year, 5, 24
	CALL write_variable_in_TIKL("Reminder to mail the Medicare Referral for November 20" & ELIG_year & ".")
	PF3
END IF


script_end_procedure_with_error_report("DAIL has been case noted. Please remember to send forms out of ECF and delete the PEPR.")
