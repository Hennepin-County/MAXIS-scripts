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
CALL changelog_update("08/29/2023", "Updated script to include creation of SPEC/MEMO with info on needed verifications.", "Mark Riegel, Hennepin County") '#363
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
	SNAP_followup_text = ", after which a new Combined Application Form (CAF) is required (expedited SNAP closing for postponed verification not returned)."
	'To do - where is cash_check coming from?
	IF cash_check <> 1 THEN intake_date = dateadd("d", 1, SNAP_last_REIN_date)			        'if cash is not being closed the intake date needs to be the day after the rein date
Else
	progs_closed = progs_closed & "SNAP/"															'if rec'd after the 15th client gets until closure date (end of 2nd month of benefits) to be reinstated.
	SNAP_last_REIN_date = closure_date
	SNAP_followup_text = ", after which a new Combined Application Form (CAF) is required (expedited SNAP closing for postponed verification not returned)."
END IF

'Navigates back to DAIL
PF3 						

'Search CASE/NOTEs to determine if there is a VERIFICATIONS REQUESTED CASE/NOTE
call write_value_and_transmit("N", 6, 3)

'To do - remove after testing complete. MAXIS case number identified through DAIL scrubber but not during testing so added in identification of MAXIS case number here
EMReadScreen MAXIS_case_number, 8, 20, 38
MAXIS_case_number = trim(replace(MAXIS_case_number,"_", " "))

'Using SNAP application date to set too old date, no need to read for dates prior to SNAP application
too_old_date = DateAdd("D", -1, SNAP_application_date)

note_row = 5
Do
	EMReadScreen note_date, 8, note_row, 6                  'reading the note date

	EMReadScreen note_title, 55, note_row, 25               'reading the note header
	note_title = trim(note_title)

	'VERIFICATIONS NOTES
	If left(note_title, 23) = "VERIFICATIONS REQUESTED" Then
		verifications_requested_case_note_found = True
		verifs_needed = ""

		EMWriteScreen "X", note_row, 3                          'Opening the VERIF note to read the verifications
		transmit

		EMReadScreen in_correct_note, 23, 4, 3                  'making sure we are in the right note
		EMReadScreen note_list_header, 23, 4, 25

		'Here we find the right row to start reading
		If in_correct_note = "VERIFICATIONS REQUESTED" Then                     'making sure we're in the right note
			in_note_row = 5
			Do
				EMReadScreen whole_note_line, 77, in_note_row, 3                'reading the whole line of the note'
				whole_note_line = trim(whole_note_line)

				in_note_row = in_note_row + 1
				If whole_note_line = "" then Exit Do
			Loop until whole_note_line = "List of all verifications requested:" 'This is the header within the note - the NEXT line starts the list of verifs

			If whole_note_line = "List of all verifications requested:" Then    'If we actually found the header.
				verif_note_lines = ""                                           'defaulting a variable to save all the lines of the note
				Do
					EMReadScreen verif_line, 77, in_note_row, 3                 'reading the line of the note
					verif_line = trim(verif_line)
					If verif_line = "" then Exit Do                             'If they are blank - we stop'
					verif_note_lines = verif_note_lines & "~|~" & verif_line    'Adding it to a string of all the lines

					in_note_row = in_note_row + 1                               'next line'

					EMReadScreen next_line, 77, in_note_row, 3                  'Looking to see if the next line is the divider line
					next_line = trim(next_line)
				Loop until next_line = "---"                                    'stop at the dividing line
				'if there were lines saved
				If verif_note_lines <> "" Then
					verif_counter = 1                                           'setting a counter to find verifs that have been numbered
					If left(verif_note_lines, 3) = "~|~" Then verif_note_lines = right(verif_note_lines, len(verif_note_lines) - 3)             'making an array of all of the lines
					If InStr(verif_note_lines, "~|~") = 0 Then
						verif_lines_array = Array(verif_note_lines)
					Else
						verif_lines_array = split(verif_note_lines, "~|~")
					End If

					verifs_to_add = ""                                          'blanking a string for adding all the lines together
					For each line in verif_lines_array
						counter_string = verif_counter & ". "                   'using the counter - which is a number to make a string that looks like what is in the note
						If left(line, 2) = "- " OR left(line, 3) = counter_string Then                          'If the string starts with a dash or the counter
							If left(line, 2) = "- " Then line = "; " & right(line, len(line) - 2)               'Removes the list delimiter and adds the editbox delimiter
							If left(line, 3) = counter_string Then line = "; " & right(line, len(line) - 3)
							verif_counter = verif_counter + 1                                                   'incrementing the counter
						Else
							line = " " & line                                                                   'adding a space to the sting so there is a space between words if we are at a 'same line'
						End If

						verifs_to_add = verifs_to_add & line                    'adding the verif information all together
					Next
					If left(verifs_to_add, 2) = "; " Then verifs_to_add = right(verifs_to_add, len(verifs_to_add) - 2)  'trimming the string
					If verifs_to_add <> "" Then verifs_needed = trim(verifs_needed & verifs_to_add) 'adding the information to the variable used in this script 
				End If
			End If
			PF3         'leaving the note
		Else
			If note_list_header <> "First line of Case note" Then PF3           'this backs us out of the note if we ended up in the wrong note.
		End If
	End If

	if note_date = "        " then Exit Do                                      'if we are at the end of the list of notes - we can't read any more

	note_row = note_row + 1
	if note_row = 19 then
		note_row = 5
		PF8
		EMReadScreen check_for_last_page, 9, 24, 14
		If check_for_last_page = "LAST PAGE" Then Exit Do
	End If
	EMReadScreen next_note_date, 8, note_row, 6
	if next_note_date = "        " then Exit Do
Loop until DateDiff("d", too_old_date, next_note_date) <= 0

'Navigate back to DAIL
PF3

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
		IF trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

'Navigates to SPEC/MEMO 
Call write_value_and_transmit("P", 6, 3)
Call write_value_and_transmit("MEMO", 20, 70)
'Function to create new MEMO
Call start_a_new_spec_memo(memo_opened, False, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, False)

'Write information to SPEC/MEMO
memo_header = "--- SNAP Closed " & closure_date & " - Expedited Autoclose ---"
Call write_variable_in_SPEC_MEMO(memo_header)
Call write_variable_in_SPEC_MEMO(" ")
Call write_variable_in_SPEC_MEMO("Reason for closure: Delayed verifications were not submitted and Expedited SNAP Autoclosed.")
Call write_variable_in_SPEC_MEMO(" ")
If verifs_needed <> "" then call write_variable_in_SPEC_MEMO("Verifications needed: " & verifs_needed)
Call write_variable_in_SPEC_MEMO(" ")
CALL write_variable_in_SPEC_MEMO("*** Submitting Documents:")
CALL write_variable_in_SPEC_MEMO("- Online at infokeep.hennepin.us or MNBenefits.mn.gov")
CALL write_variable_in_SPEC_MEMO("  Use InfoKeep to upload documents directly to your case.")
CALL write_variable_in_SPEC_MEMO("- Mail, Fax, or Drop Boxes at service centers listed below:")
CALL write_variable_in_SPEC_MEMO(" - 7051 Brooklyn Blvd Brooklyn Center 55429")
CALL write_variable_in_SPEC_MEMO(" - 1011 1st St S Hopkins 55343")
CALL write_variable_in_SPEC_MEMO(" - 1001 Plymouth Ave N Minneapolis 55411")
CALL write_variable_in_SPEC_MEMO(" - 2215 East Lake Street Minneapolis 55407")
CALL write_variable_in_SPEC_MEMO(" (Hours are 8 - 4:30 Monday - Friday)")
Call write_variable_in_SPEC_MEMO(" ")
'To do - need to verify if this is necessary, won't it always include this info? 
If case_noting_intake_dates = True or case_noting_intake_dates = "" then call write_variable_in_SPEC_MEMO("Last SNAP reinstatement date: " & SNAP_last_REIN_date & SNAP_followup_text)
'Save MEMO and navigate back to DAIL
PF4
PF3

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
If verifs_needed <> "" then call write_bullet_and_variable_in_case_note("Verifications needed", verifs_needed)

'To do - check on the case noting intake notes
If case_noting_intake_dates = True or case_noting_intake_dates = "" then call write_bullet_and_variable_in_case_note("Last SNAP REIN date", SNAP_last_REIN_date & SNAP_followup_text)
call write_variable_in_case_note("---")
call write_variable_in_case_note(worker_signature)

script_end_procedure("")
