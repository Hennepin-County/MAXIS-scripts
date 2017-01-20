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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.

EMReadscreen MAXIS_case_number, 8, 5, 72			'Reading the case number and footer month/year
EMReadscreen MAXIS_footer_month, 2, 6, 11
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

PF3 						'Navigates to Case Notes from DAIL - maintaining the tie to the list - does this so worker can reveiw case notes while the next dialog is up
EMWriteScreen "N", 6, 3
transmit

'Dialog is defined here so the case number and application date are listed on it
BeginDialog verifs_dialog, 0, 0, 401, 70, "Verifications Needed"
  EditBox 90, 5, 305, 15, verifs_needed
  EditBox 305, 30, 90, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 295, 50, 50, 15
    CancelButton 345, 50, 50, 15
  Text 235, 35, 65, 10, "Sign your case note"
  Text 10, 10, 70, 10, "Verifications Needed:"
  Text 10, 35, 65, 10, "Date of application"
  Text 95, 35, 90, 10, SNAP_application_date
  Text 10, 50, 50, 10, "Case Number"
  Text 95, 50, 75, 10, MAXIS_case_number
EndDialog

'Runs the dialog to allow workers to sign and to list verifications needed
Do
	Dialog verifs_dialog
	If ButtonPressed = Cancel Then StopScript
Loop until ButtonPressed = OK

'The script checks to make sure it is on the NOTES main list
EMReadScreen notes_check, 9, 2, 33
EMReadScreen mode_check, 1, 20, 09
If notes_check = "Case Note" AND mode_check = " " Then
	PF9					'If so, it starts a new case note 'and maintains the tie to the DAIL list
Else
	start_a_blank_CASE_NOTE		'If the worker navigated away from NOTES, this will get a new case note started but will not maintain the tie to the DAIL list
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
