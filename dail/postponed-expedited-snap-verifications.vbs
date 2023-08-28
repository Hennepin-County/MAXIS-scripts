'Required for statistical purposes===============================================================================
name_of_script = "DAIL - POSTPONED XFS VERIFICATIONS.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 64          'manual run time in seconds
STATS_denomination = "C"       'C is for Case
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
CALL changelog_update("09/16/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.
EMReadscreen MAXIS_footer_month, 2, 6, 11               	'Reading the footer month/year
EMReadscreen MAXIS_footer_year, 2, 6, 14

dail_date = MAXIS_footer_month & "/01/" & MAXIS_footer_year			'Creates a date out of the footer month and year
closure_date = dateadd("d", -1, dail_date)							'Closure date is always one day before the month the DAIL message was issued

EMWriteScreen "S", 6, 3							'Navigates to STAT PROG - function not used to maintain tie to DAIL list
transmit
EMWriteScreen "PROG", 20, 71
transmit
EMReadScreen SNAP_application_date, 8, 10, 33	'Reads the date of application for SNAP and formats it
SNAP_application_date = replace(SNAP_application_date, " ", "/")
fifteen_of_appl_month = left(SNAP_application_date, 2) & "/15/" & right(SNAP_application_date, 2)
IF datediff("D", SNAP_application_date, fifteen_of_appl_month) >= 0 Then							'if rec'd ON or BEFORE 15th client gets 30 days from date of application to be reinstated.
	progs_closed = progs_closed & "SNAP/"
	SNAP_last_REIN_date = dateadd("d", 30, SNAP_application_date)
	SNAP_followup_text = ", after which a new CAF is required (expedited SNAP closing for postponed verification not returned)."
	IF cash_check <> 1 THEN intake_date = dateadd("d", 1, SNAP_last_REIN_date)			        'if cash is not being closed the intake date needs to be the day after the rein date
Else
	progs_closed = progs_closed & "SNAP/"															'if rec'd after the 15th client gets until closure date (end of 2nd month of benefits) to be reinstated.
	SNAP_last_REIN_date = closure_date
	SNAP_followup_text = ", after which a new CAF is required (expedited SNAP closing for postponed verification not returned)."
END IF

'Navigates back to DAIL
PF3 						

'Navigates to SPEC/MEMO through DAIL to maintain tie to list
write_value_and_transmit("P", 6, 3)
write_value_and_transmit("MEMO", 20, 70)
'To do - double-check function(ality), ran into issues leaving the DAIL
' Call start_a_new_spec_memo(memo_opened, search_for_arep_and_swkr, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, False)
'Opens new MEMO
PF5
'Creates new MEMO
write_value_and_transmit("X", 5, 12)

'Dialog is defined here so the case number and application date are listed on it
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 401, 120, "Verifications Needed"
	Text 5, 5, 390, 20, "Please review and confirm the required verifications below. Utilize plain language for the verifications as this will be used to create a MEMO and CASE/NOTE."
	Text 10, 35, 70, 10, "Verifications Needed:"
  	EditBox 85, 30, 305, 15, verifs_needed
  	Text 10, 50, 65, 10, "Date of application:"
  	Text 85, 50, 90, 10, SNAP_application_date
  	Text 10, 65, 50, 10, "Case Number:"
  	Text 85, 65, 75, 10, MAXIS_case_number
	Text 10, 85, 65, 10, "Worker Signature:"
  	EditBox 85, 80, 90, 15, worker_signature
  	ButtonGroup ButtonPressed
    	OkButton 290, 100, 50, 15
    	CancelButton 345, 100, 50, 15
EndDialog

'Runs the dialog to allow workers to sign and to list verifications needed
DO
	DO
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation
		Call validate_MAXIS_case_number(err_msg, "*")
		IF trim(verifs_needed) = "" THEN err_msg = err_msg & vbCr & "* Please enter the postponed verifications requested."
		IF len(trim(verifs_needed)) > 30 THEN err_msg = err_msg & vbCr & "* The list of verifications must be less than 30 characters long. Please shorten."
		IF trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

'To do - update navigation back to/from DAIL to CASE/NOTE
'Navigates to Case Notes from DAIL - maintaining the tie to the list - does this so worker can reveiw case notes while the next dialog is up
EMWriteScreen "N", 6, 3
transmit

'The script checks to make sure it is on the NOTES main list
EMReadScreen notes_check, 9, 2, 33
EMReadScreen mode_check, 1, 20, 09
If notes_check = "Case Note" AND mode_check = " " Then
	PF9					'If so, it starts a new case note 'and maintains the tie to the DAIL list
Else
    Call navigate_to_MAXIS_screen("CASE", "NOTE")
    PF9 'edit mode
End If
'This is the CASE NOTE --------------------------------------------------------------------'
case_note_header = "--- SNAP Closed " & closure_date & " - Expedited Autoclose ---"
call write_variable_in_case_note(case_note_header)
call write_bullet_and_variable_in_case_note("Reason for closure", "Delayed verifications were not submitted and Expedited SNAP Autoclosed.")
If verifs_needed <> "" then call write_bullet_and_variable_in_case_note("Verifs needed", verifs_needed)
If case_noting_intake_dates = True or case_noting_intake_dates = "" then call write_bullet_and_variable_in_case_note("Last SNAP REIN date", SNAP_last_REIN_date & SNAP_followup_text)
call write_variable_in_case_note("---")
call write_variable_in_case_note(worker_signature)

script_end_procedure("")
