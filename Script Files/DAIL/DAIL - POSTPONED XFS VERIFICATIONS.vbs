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


'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.

EMReadscreen MAXIS_case_number, 8, 5, 72
EMReadscreen DAIL_month, 2, 6, 11
EMReadscreen DAIL_year, 2, 6, 13

EMWriteScreen "S", 6, 3
transmit
EMWriteScreen "PROG", 20, 71
transmit
EMReadScreen SNAP_application_date, 8, 10, 33
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
	IF cash_check <> 1 THEN intake_date = dateadd("d", 1, SNAP_last_REIN_date)					'if cash is not being closed the intake date needs to be the day after the rein date
END IF

'The dialog navigated to CASE/NOTE. This will write the info into the case note.
start_a_blank_CASE_NOTE
IF death_check = 1 THEN
	case_note_header = "---Closed " & progs_closed & " due to client death---"
ELSE
	case_note_header = "---Closed " & progs_closed & " " & closure_date & "---"
END IF
call write_variable_in_case_note(case_note_header)
call write_bullet_and_variable_in_case_note("Reason for closure", reason_for_closure)
If verifs_needed <> "" then call write_bullet_and_variable_in_case_note("Verifs needed", verifs_needed)
If WCOM_check = 1 then call write_variable_in_case_note("* Added WCOM to notice.")
If case_noting_intake_dates = True or case_noting_intake_dates = "" then  																								'Updated bug. With new/updated, scripts this functionality was not being accessed.'
	call write_bullet_and_variable_in_case_note("Last SNAP REIN date", SNAP_last_REIN_date & SNAP_followup_text)
	If death_check <> 1 and cash_check = 1 then call write_bullet_and_variable_in_case_note("Last cash REIN date", cash_last_REIN_date & cash_followup_text)
''	If open_progs <> "" and len(open_progs) > 1 and open_progs <> "no" and open_progs <> "NO" and open_progs <> "No" and open_progs <> "n/a" and open_progs <> "N/A" and open_progs <> "NA" and open_progs <> "na" then
''		call write_bullet_and_variable_in_case_note("Open programs", open_progs)
''	Else
''		IF death_check = 1 THEN call write_variable_in_case_note("* All programs closed.")
''		IF death_check <> 1 THEN call write_variable_in_case_note("* All programs closed. Cash and/or SNAP (whichever applicable) case becomes intake again on " & intake_date & ".")
''	End if
Else
	If open_progs <> "" and len(open_progs) > 1 and open_progs <> "no" and open_progs <> "NO" and open_progs <> "No" and open_progs <> "n/a" and open_progs <> "N/A" and open_progs <> "NA" and open_progs <> "na" then call write_bullet_and_variable_in_case_note("Open programs", open_progs)
End if

call write_variable_in_case_note("---")
call write_variable_in_case_note(worker_signature)


