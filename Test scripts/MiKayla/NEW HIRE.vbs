'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.

'Required for statistical purposes===============================================================================
name_of_script = "DAIL - ALL NEW HIRE.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 345         'manual run time in seconds
STATS_denomination = "C"       'C is for each MEMBER
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
call changelog_update("08/16/2019", "Initial version.", "MiKayla Handley, Hennepin County")'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'CONNECTS TO DEFAULT SCREEN
EMConnect ""
match_found = FALSE
'CHECKS TO MAKE SURE THE WORKER IS ON THEIR DAIL
EMReadscreen dail_check, 4, 2, 48
If dail_check <> "DAIL" then script_end_procedure("You are not in your dail. This script will stop.")

EMGetCursor dail_row, dail_col

'TYPES A "T" TO BRING THE SELECTED MESSAGE TO THE TOP
EMSendKey "t"
TRANSMIT

'New HIRE messages, client started a new job (loads NEW HIRE)
EMReadScreen HIRE_check, 15, 6, 20
If HIRE_check = "NEW JOB DETAILS" then
    match_found = TRUE
	SDNH_match_type = TRUE
	'call run_from_GitHub(script_repository & "dail/new-hire.vbs")
END IF
'New HIRE messages, client started a new job (loads NEW HIRE)
EMReadScreen HIRE_check, 11, 6, 37
If HIRE_check = "JOB DETAILS" then
	match_found = TRUE
	NDNH_match_type = TRUE
    'call run_from_GitHub(script_repository & "dail/new-hire-ndnh.vbs")
END IF


'The following reads the message in full for the end part (which tells the worker which message was selected)
EMReadScreen full_message, 60, 6, 20
full_message = trim(full_message)
EmReadScreen MAXIS_case_number, 8, 5, 73
MAXIS_case_number = trim(MAXIS_case_number)

EMReadScreen extra_info, 1, 6, 80
IF extra_info = "+" or extra_info = "&" THEN
	EMSendKey "X"
	TRANSMIT
	'THE ENTIRE MESSAGE TEXT IS DISPLAYED'
	EmReadScreen error_msg, 37, 24, 02
	row = 1
	col = 1
	EMSearch "Case Number", row, col 	'Has to search, because every once in a while the rows and columns can slide one or two positions.
	'If row = 0 then script_end_procedure("MAXIS may be busy: the script appears to have errored out. This should be temporary. Try again in a moment. If it happens repeatedly contact the alpha user for your agency.")
	EMReadScreen first_line, 61, row + 3, col - 40 'JOB DETAIL Reads each line for the case note. COL needs to be subtracted from because of NDNH message format differs from original new hire format.
		'first_line = replace(first_line, "FOR  ", "FOR ")	'need to replaces 2 blank spaces'
		first_line = trim(first_line)
	EMReadScreen second_line, 61, row + 4, col - 40
		second_line = trim(second_line)
	EMReadScreen third_line, 61, row + 5, col - 40 'maxis name'
		third_line = trim(third_line)
		'third_line = replace(third_line, ",", ", ")
	EMReadScreen fourth_line, 61, row + 6, col - 40'new hire name'
		fourth_line = trim(fourth_line)
		'fourth_line = replace(fourth_line, ",", ", ")
	EMReadScreen fifth_line, 61, row + 7, col - 40'new hire name'
		fifth_line = trim(fifth_line)
	TRANSMIT
END IF

'The script needs to determine what the day is in a MAXIS friendly format. The following does that.
current_month = CM_mo
If len(current_month) = 1 then current_month = "0" & current_month
current_day = datepart("d", date)
If len(current_day) = 1 then current_day = "0" & current_day
current_year = CM_yr
'SELECTS THE DAIL MESSAGE AND READS THE RESPONSE
EMSendKey "x"
TRANSMIT
row = 1
col = 1
msgbox NDNH_match_type & " or " & SDNH_match_type
IF NDNH_match_type = TRUE THEN
    EMSearch "JOB DETAILS", row, col 	'Has to search, because every once in a while the rows and columns can slide one or two positions.
    If row = 0 then script_end_procedure_with_error_report("MAXIS may be busy: the script appears to have errored out. This should be temporary. Try again in a moment. If it happens repeatedly contact the alpha user for your agency.")
    EMReadScreen new_hire_first_line, 61, row, col - 7 'JOB DETAIL Reads each line for the case note. COL needs to be subtracted from because of NDNH message format differs from original new hire format.
    	new_hire_first_line = replace(new_hire_first_line, "FOR  ", "FOR ")	'need to replaces 2 blank spaces'
    	new_hire_first_line = trim(new_hire_first_line)
    EMReadScreen new_hire_second_line, 61, row + 1, col -15
    	new_hire_second_line = trim(new_hire_second_line)
    EMReadScreen new_hire_third_line, 61, row + 2, col -15 'maxis name'
    	new_hire_third_line = trim(new_hire_third_line)
    	new_hire_third_line = replace(new_hire_third_line, ",", ", ")
    EMReadScreen new_hire_fourth_line, 61, row + 3, col -15'new hire name'
    	new_hire_fourth_line = trim(new_hire_fourth_line)
    	new_hire_fourth_line = replace(new_hire_fourth_line, ",", ", ")
    'IF right(new_hire_third_line, 46) <> right(new_hire_fourth_line, 46) then 				'script was being run on cases where the names did not match but SSN did. This will allow users to review.
    '	warning_box = MsgBox("The names found on the NEW HIRE message do not match exactly." & vbcr & new_hire_third_line & vbcr & new_hire_fourth_line & vbcr & "Please review and click OK if you wish to continue and CANCEL if the name is incorrect.", vbOKCancel)
    '	If warning_box = vbCancel then script_end_procedure("The script has ended. Please review the new hire as you indicated that the name read from the NEW HIRE and the MAXIS name did not match.")
    'END IF
    row = 1 						'Now it's searching for info on the hire date as well as employer
    col = 1
    'Now it's searching for info on the hire date as well as employer
    EMSearch "DATE HIRED", row, col
    EMReadScreen date_hired, 10, row, col + 15
    If date_hired = "  -  -  EM" OR date_hired = "UNKNOWN  E" then date_hired = current_month & "-" & current_day & "-" & current_year
    date_hired = trim(date_hired)
    'date_hired = CDate(date_hired)
    month_hired = Datepart("m", date_hired)
    If len(month_hired) = 1 then month_hired = "0" & month_hired
    day_hired = Datepart("d", date_hired)
    If len(day_hired) = 1 then day_hired = "0" & day_hired
    year_hired = Datepart("yyyy", date_hired)
    year_hired = year_hired - 2000
    EMSearch "EMPLOYER:", row, col
    EMReadScreen employer, 25, row, col + 10
    employer = TRIM(employer)
    EMReadScreen new_HIRE_SSN, 9, 9, 5
    PF3
END IF
IF SDNH_match_type = TRUE THEN
    EMSearch "NEW JOB DETAILS", row, col 	'Has to search, because every once in a while the rows and columns can slide one or two positions.
    If row = 0 then script_end_procedure_with_error_report("MAXIS may be busy: the script appears to have errored out. This should be temporary. Try again in a moment. If it happens repeatedly contact the alpha user for your agency.")
    EMReadScreen new_hire_first_line, 61, row, col'JOB DETAIL Reads each line for the case note. COL needs to be subtracted from because of NDNH message format differs from original new hire format.
    	new_hire_first_line = replace(new_hire_first_line, "FOR  ", "FOR ")	'need to replaces 2 blank spaces'
    	new_hire_first_line = trim(new_hire_first_line)
    EMReadScreen new_hire_second_line, 61, row + 1, col
    	new_hire_second_line = trim(new_hire_second_line)
    EMReadScreen new_hire_third_line, 61, row + 2, col 'maxis name'
    	new_hire_third_line = trim(new_hire_third_line)
    	new_hire_third_line = replace(new_hire_third_line, ",", ", ")
    EMReadScreen new_hire_fourth_line, 61, row + 3, col'new hire name'
    	new_hire_fourth_line = trim(new_hire_fourth_line)
    	new_hire_fourth_line = replace(new_hire_fourth_line, ",", ", ")
    IF right(new_hire_third_line, 46) <> right(new_hire_fourth_line, 46) then 				'script was being run on cases where the names did not match but SSN did. This will allow users to review.
    	warning_box = MsgBox("The names found on the NEW HIRE message do not match exactly." & vbcr & new_hire_third_line & vbcr & new_hire_fourth_line & vbcr & "Please review and click OK if you wish to continue and CANCEL if the name is incorrect.", vbOKCancel)
    	If warning_box = vbCancel then script_end_procedure_with_error_report("The script has ended. Please review the new hire as you indicated that the name read from the NEW HIRE and the MAXIS name did not match.")
    END IF
    row = 1 						'Now it's searching for info on the hire date as well as employer
    col = 1
    'Now it's searching for info on the hire date as well as employer
    EMSearch "DATE HIRED:", row, col
    EMReadScreen date_hired, 10, row, col + 12
    If date_hired = "  -  -  EM" OR date_hired = "UNKNOWN  E" then date_hired = current_month & "-" & current_day & "-" & current_year
    date_hired = trim(date_hired)
    'date_hired = CDate(date_hired)'
    month_hired = Datepart("m", date_hired)
    If len(month_hired) = 1 then month_hired = "0" & month_hired
    day_hired = Datepart("d", date_hired)
    If len(day_hired) = 1 then day_hired = "0" & day_hired
    year_hired = Datepart("yyyy", date_hired)
    year_hired = year_hired - 2000
    EMSearch "EMPLOYER:", row, col
    EMReadScreen employer, 25, row, col + 10
    employer = TRIM(employer)

    row = 1 						'Now it's searching for the SSN
    col = 1
    EMSearch "SSN #", row, col
    EMReadScreen new_HIRE_SSN, 11, row, col + 5
    new_HIRE_SSN = TRIM(new_HIRE_SSN)
    new_HIRE_SSN = replace(new_HIRE_SSN, "-", "") '01/10/19 it has the dashes in the match now'
    PF3
END IF

BeginDialog NDNH_only_dialog, 0, 0, 236, 70, "New Hires Action"
  DropListBox 150, 5, 80, 15, "Select One:"+chr(9)+"NO-RUN NEW HIRE"+chr(9)+"YES-INFC clear match", match_answer_droplist
  ButtonGroup ButtonPressed
    OkButton 125, 50, 50, 15
    CancelButton 180, 50, 50, 15
  Text 10, 10, 140, 10, "Has this match been acted on previously?"
  Text 30, 25, 190, 20, "Reminder that client must be provided 10 days to return                             requested verification(s)"
EndDialog


Do
	DO
		err_msg = ""
		Dialog NDNH_only_dialog
		cancel_confirmation
		IF match_answer_droplist = "Select One:" THEN MsgBox("You must select an answer.")
	LOOP UNTIL match_answer_droplist <> "Select One:"
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
LOOP UNTIL err_msg = ""

If match_answer_droplist = "NO-RUN NEW HIRE" THEN
    'CHECKING CASE CURR. MFIP AND SNAP HAVE DIFFERENT RULES.
    EMWriteScreen "h", 6, 3
    TRANSMIT
    row = 1
    col = 1
    EMSearch "FS: ", row, col
    If row <> 0 then FS_case = True
    If row = 0 then FS_case = False
    row = 1
    col = 1
    EMSearch "MFIP: ", row, col
    If row <> 0 then MFIP_case = True
    If row = 0 then MFIP_case = False
    PF3
    If dail_row <> 6 Then Call write_value_and_transmit("t", dail_row, 3)       'bringing the correct message back to the top'
    'GOING TO STAT
    EMSendKey "s"
    TRANSMIT
    EMReadScreen stat_check, 4, 20, 21
    If stat_check <> "STAT" then script_end_procedure_with_error_report("Unable to get to stat due to an error screen. Clear the error screen and return to the DAIL. Then try the script again.")
    'GOING TO MEMB, NEED TO CHECK THE HH MEMBER
    EMWriteScreen "memb", 20, 71
    TRANSMIT
    Do
    	EMReadScreen MEMB_current, 1, 2, 73
    	EMReadScreen MEMB_total, 1, 2, 78
    	EMReadScreen MEMB_SSN, 11, 7, 42
    	If new_HIRE_SSN = replace(MEMB_SSN, " ", "") then
    		EMReadScreen HH_memb, 2, 4, 33
    		EMReadScreen memb_age, 2, 8, 76
    		If cint(memb_age) < 19 then MsgBox "This client is under 19, so make sure to check that school verification is on file."
    	End if
    		TRANSMIT
    LOOP UNTIL (MEMB_current = MEMB_total) or (new_HIRE_SSN = replace(MEMB_SSN, " ", ""))
    		'GOING TO JOBS
    EMWriteScreen "jobs", 20, 71
    EMWriteScreen HH_memb, 20, 76
    TRANSMIT
    'MFIP cases need to manually add the JOBS panel for ES purposes.
    If MFIP_case = False then create_JOBS_checkbox = checked
    'Defaulting the "set TIKL" variable to checked
    TIKL_checkbox = CHECKED
    'Setting the variable for the following do...loop
    HH_memb_row = 5
    'Show dialog

    'DIALOGS----------------------------------------------------------------------------------------------

    'This is a dialog asking if the job is known to the agency.
    BeginDialog new_HIRE_dialog, 0, 0, 281, 195, "NEW HIRE INFORMATION"
      EditBox 65, 5, 20, 15, HH_memb
      EditBox 65, 25, 95, 15, employer
      CheckBox 15, 45, 190, 10, "Check here to have the script create a new JOBS panel.", create_JOBS_checkbox
      CheckBox 15, 55, 160, 10, "Job is known to the agency (exit the script).", job_known_checkbox
      CheckBox 15, 80, 195, 10, "Sent a request for verifications out of ECF.", TIKL_checkbox
      CheckBox 15, 90, 190, 10, "Sent a Work Number request and submitted to ECF. ", work_number_checkbox
      CheckBox 15, 100, 100, 10, "Requesting CEI/OHI docs.", requested_CEI_OHI_docs_checkbox
      CheckBox 15, 110, 105, 10, "Sent a status update to CCA.", CCA_checkbox
      CheckBox 15, 120, 100, 10, "Sent a status update to ES. ", ES_checkbox
      CheckBox 15, 140, 135, 10, "Check to have an outlook reminder set", Outlook_reminder_checkbox
      EditBox 65, 155, 210, 15, other_notes
      EditBox 65, 175, 110, 15, worker_signature
      ButtonGroup ButtonPressed
        OkButton 190, 175, 40, 15
        CancelButton 235, 175, 40, 15
        PushButton 175, 15, 45, 10, "prev. panel", prev_panel_button
        PushButton 175, 25, 45, 10, "next panel", next_panel_button
        PushButton 225, 15, 45, 10, "prev. memb", prev_memb_button
        PushButton 225, 25, 45, 10, "next memb", next_memb_button
      Text 5, 180, 60, 10, "Worker signature:"
      GroupBox 170, 5, 105, 35, "STAT-based navigation"
      Text 5, 30, 50, 10, "New Hire Info:"
      Text 20, 160, 40, 10, "Other notes:"
      Text 5, 10, 60, 10, "Member Number:"
      GroupBox 5, 70, 270, 65, "Verification or updates"
    EndDialog


    DO
        DO
        	Dialog new_HIRE_dialog
        	cancel_confirmation
        	MAXIS_dialog_navigation
        LOOP UNTIL ButtonPressed = -1
        call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
    LOOP UNTIL are_we_passworded_out = false
    EMWriteScreen "jobs", 20, 71
    EMWriteScreen HH_memb, 20, 76
    TRANSMIT
    'Checking to see if 5 jobs already exist. If so worker will need to manually delete one first.
    EMReadScreen jobs_total_panel_count, 1, 2, 78
    IF create_JOBS_checkbox = checked AND jobs_total_panel_count = "5" THEN script_end_procedure_with_error_report("This client has 5 jobs panels already. Please review and delete and unneeded panels if you want the script to add a new one.")

    'If new job is known, script ends.
    If job_known_checkbox = checked then script_end_procedure("The script will stop as this job is known.")
    'Now it will create a new JOBS panel for this case.
    If create_JOBS_checkbox = checked then
    	EMWriteScreen "nn", 20, 79				'Creates new panel
    	TRANSMIT
    	EMReadScreen MAXIS_footer_month, 2, 20, 55	'Reads footer month for updating the panel
    	EMReadScreen MAXIS_footer_year, 2, 20, 58		'Reads footer year
    	EMWriteScreen "w", 5, 34				'Wage income is the type
    	EMWriteScreen "n", 6, 34				'No proof has been provided
    	EMWriteScreen employer, 7, 42			'Adds employer info
    	EMWriteScreen month_hired, 9, 35		'Adds month hired to start date (this is actually the day income was received)
    	EMWriteScreen day_hired, 9, 38			'Adds day hired
    	EMWriteScreen year_hired, 9, 41			'Adds year hired
    	EMWriteScreen MAXIS_footer_month, 12, 54		'Puts footer month in as the month on prospective side of panel
      	IF month_hired = MAXIS_footer_month THEN     'This accounts for rare cases when new hire footer month is the same as the hire date.
      		EMWriteScreen day_hired, 12, 57			'Puts date hired if message is from same month as hire ex 01/16 new hire for 1/17/16 start date.
      	ELSE
      		EMWriteScreen "01", 12, 57		'Puts the first in as the day on prospective side
      	END IF
      	EMWriteScreen MAXIS_footer_year, 12, 60		'Puts footer year in on prospective side
      	EMWriteScreen "0", 12, 67				'Puts $0 in as the received income amt
      	EMWriteScreen "0", 18, 72				'Puts 0 hours in as the worked hours
      	If FS_case = True then 					'If case is SNAP, it creates a PIC
      		EMWriteScreen "x", 19, 38
      		TRANSMIT
      		IF month_hired = MAXIS_footer_month THEN     'This accounts for rare cases when new hire footer month is the same as the hire date.
      			EMWriteScreen month_hired, 5, 34
      			EMWriteScreen day_hired, 5, 37
      			EMWriteScreen year_hired, 5, 40
      		ELSE
      		EMWriteScreen current_month, 5, 34
        	EMWriteScreen current_day, 5, 37
        	EMWriteScreen current_year, 5, 40
      		END IF
      		EMWriteScreen "1", 5, 64
      		EMWriteScreen "0", 8, 64
      		EMWriteScreen "0", 9, 66
      		TRANSMIT
      		TRANSMIT
      		TRANSMIT
    	END IF
    	TRANSMIT						'Transmits to submit the panel
      	EMReadScreen expired_check, 6, 24, 17 'Checks to see if the jobs panel will carry over by looking for the "This information will expire" at the bottom of the page
      	If expired_check = "EXPIRE" THEN Msgbox "Check next footer month to make sure the JOBS panel carried over"
    END IF
    PF3 'back to DAIL
    'TRANSMIT
    '-----------------------------------------------------------------------------------------CASENOTE
    start_a_blank_case_note	'Writes that the message is unreported, and that the proofs are being sent/TIKLed for.
    IF NDNH_match_type = TRUE THEN CALL write_variable_in_case_note("-NDNH JOB DETAILS FOR (M" & HH_memb & ") unreported to agency-")
    IF SDNH_match_type = TRUE THEN CALL write_variable_in_case_note("-SDNH JOB DETAILS FOR (M" & HH_memb & ") unreported to agency-")
    CALL write_variable_in_case_note("DATE HIRED: " & date_hired)
    CALL write_variable_in_case_note("EMPLOYER: " & employer)
    CALL write_variable_in_case_note(new_hire_third_line)
    CALL write_variable_in_case_note(new_hire_fourth_line)
    CALL write_variable_in_case_note("---")
    IF TIKL_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Sent employment verification and DHS-2919B (Verif Request Form B) from ECF & TIKLed for 10-day return. ")
    IF create_JOBS_checkbox = checked THEN CALL write_variable_in_case_note("* STAT/JOBS updated with new hire information from DAIL.")
    IF CCA_checkbox = CHECKED  THEN CALL write_variable_in_case_note("* Sent status update to CCA.")
    IF ES_checkbox = CHECKED  THEN CALL write_variable_in_case_note("* Sent status update to ES.")
    IF work_number_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Sent request for Work Number after confirming client authorization.")
    IF CEI_checkbox = checked THEN CALL write_variable_in_case_note("* Requested CEI/OHI docs.")
    CALL write_bullet_and_variable_in_case_note("Other notes", other_notes)
    CALL write_variable_in_case_note("---")
    CALL write_variable_in_case_note(worker_signature)
    PF3
        'If TIKL_checkbox is unchecked, it needs to end here.
    IF TIKL_checkbox = UNCHECKED THEN script_end_procedure_with_error_report("Success! MAXIS updated for new NDNH HIRE message, and a case note made. An Employment Verification and Verif Req Form B should now be sent. The job is at " & employer & ".")
    'Navigates to TIKL
    IF TIKL_checkbox = CHECKED THEN
    	Call navigate_to_MAXIS_screen("DAIL", "WRIT")
    	CALL create_MAXIS_friendly_date(date, 10, 5, 18)   'The following will generate a TIKL formatted date for 10 days from now, and add it to the TIKL
    	CALL write_variable_in_TIKL("Verification of " & employer & "job via NEW HIRE should have returned by now. If not received and processed, take appropriate action." & vbcr & "For all federal matches INFC/HIRE must be cleared please see HSR manual.")
    	PF3		'Exits and saves TIKL
    	script_end_procedure_with_error_report("Success! MAXIS updated for new HIRE message, a case note made, and a TIKL has been sent for 10 days from now. An Employment Verification and Verif Req Form B should now be sent. The job is at " & employer & ".")
    END IF
END IF

IF match_answer_droplist = "YES-INFC clear match" THEN
'This is a dialog asking if the job is known to the agency.
    BeginDialog Match_Info_dialog, 0, 0, 281, 250, "NDNH Match Resolution Information"
      CheckBox 10, 15, 265, 10, "Check here to verify that ECF has been reviewed and acted upon appropriately", ECF_checkbox
      DropListBox 170, 35, 95, 15, "Select One:"+chr(9)+"YES-No Further Action"+chr(9)+"NO-See Next Question", Emp_known_droplist
      DropListBox 170, 55, 95, 15, "Select One:"+chr(9)+"NA-No Action Taken"+chr(9)+"BR-Benefits Reduced"+chr(9)+"CC-Case Closed", Action_taken_droplist
      EditBox 220, 75, 45, 15, cost_savings
      EditBox 55, 145, 210, 15, other_notes
      CheckBox 10, 175, 260, 10, "Check here if 10 day cutoff has passed - TIKL will be set for following month", tenday_checkbox
      CheckBox 10, 210, 185, 10, "Check to update claim referral tracking(SNAP and MF)", claim_referral_tracking_checkbox
      ButtonGroup ButtonPressed
        OkButton 170, 230, 50, 15
        CancelButton 225, 230, 50, 15
      GroupBox 5, 5, 270, 25, "ECF review"
      Text 10, 40, 145, 10, "Was this employment known to the agency?"
      Text 10, 60, 155, 10, "If unknown: what action was taken by agency?"
      Text 10, 80, 165, 10, "**First month cost savings (enter only numbers):"
      Text 5, 150, 40, 10, "Other notes:"
      GroupBox 5, 165, 270, 25, "10 day cutoff for closure"
      GroupBox 5, 200, 270, 25, "Claim Referral Tracking"
      Text 10, 95, 255, 40, "**First-month cost savings equals the difference between the monthly benefit the SNAP case would have been given in the absence of employment information resulting from an NDNH match, and the monthly benefit the SNAP case actually receive after the verified employment information was used to modify benefits. - Per TE02.08.142"
    EndDialog

	'naviagting into INFC'
	EMSendKey "I"
	TRANSMIT
	EMWriteScreen "HIRE", 20, 71
	TRANSMIT
	row = 9
	DO
        EMReadScreen case_number, 8, row, 5
        case_number = trim(case_number)
        IF case_number = MAXIS_case_number THEN
    		EMReadScreen employer_match, 20, row, 36
    		employer_match = trim(employer_match)
    		IF trim(employer_match) = "" THEN script_end_procedure("An employer match for the could not be found. The script will now end.")
    	  	IF employer_match = employer THEN
    	   		EMReadScreen cleared_value, 1, row, 61
    			IF cleared_value = " " THEN
                    EmReadscreen date_of_hire, 8, row, 20
                    EmReadscreen match_month, 5, row, 14
    				info_confirmation = MsgBox("Press YES to confirm this is the match you wish to act on." & vbNewLine & "For the next match, press NO." & vbNewLine & vbNewLine & _
    				"   " & employer_match & vbNewLine & "     Case: " & case_number & vbNewLine & "     Hire Date: " & date_of_hire & vbNewLine & "     Month: " & match_month, vbYesNoCancel, "Please confirm this match")
    				IF info_confirmation = vbCancel THEN script_end_procedure ("The script has ended. The match has not been acted on.")
    				IF info_confirmation = vbYes THEN
                        hire_match = TRUE
                        match_row = row
                        EXIT DO
                    END IF
    			END IF
    	  	END IF
        END IF
        row = row + 1
		IF row = 19 THEN
			PF8
            EmReadscreen end_of_list, 9, 24, 14
            If end_of_list = "LAST PAGE" Then Exit Do
			row = 9
		END IF
	LOOP UNTIL case_number = ""
	IF hire_match <> TRUE THEN script_end_procedure("No pending HIRE match found. Please review NEW HIRE.")
	DO
		DO
			err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
			Dialog Match_Info_dialog
			cancel_confirmation
			IF ECF_checkbox = UNCHECKED THEN err_msg = err_msg & vbCr & "* You must check that you reviewed ECF and the HIRE was acted on appropriately."
			IF Emp_known_droplist = "Select One:" THEN err_msg = err_msg & vbCr & "* You must select yes or no for was this employment known to the agency?"
			IF (Emp_known_droplist = "YES-No Further Action" AND Action_taken_droplist <> "Select One:") THEN err_msg = err_msg & vbCr & "* The employment is known and no selection needs to made for action taken."
			IF (Emp_known_droplist = "NO-See Next Question" AND Action_taken_droplist = "Select One:") THEN err_msg = err_msg & vbCr & "* You must select an action taken."
			IF (Action_taken_droplist = "NA-No Action Taken" AND cost_savings <> "") THEN err_msg = err_msg & vbCr & "* Please remove Cost savings information or make another selection"
			IF (Action_taken_droplist = "BR-Benefits Reduced" OR Action_taken_droplist = "CC-Case Closed") AND cost_savings = "" THEN err_msg = err_msg & vbCr & "* Enter the 1st month's cost savings for this case."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
		LOOP UNTIL err_msg = ""									'loops until all errors are resolved
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in
	'entering the INFC/HIRE match '
	EMWriteScreen "U", match_row, 3
	TRANSMIT
	EMReadscreen panel_check, 4, 2, 49
	IF panel_check <> "NHMD" THEN msgbox "We did not enter to clear the match"
	IF Emp_known_droplist = "NO-See Next Question" THEN EMWriteScreen "N", 16, 54
	IF Emp_known_droplist = "YES-No Further Action" THEN EMWriteScreen "Y", 16, 54
	IF Action_taken_droplist = "NA-No Action Taken" THEN EMWriteScreen "NA", 17, 54
	IF Action_taken_droplist = "BR-Benefits Reduced" THEN EMWriteScreen "BR", 17, 54
	IF Action_taken_droplist = "CC-Case Closed" THEN EMWriteScreen "CC", 17, 54
	IF cost_savings <> "" THEN
		cost_savings = round(cost_savings)
		EMWriteScreen cost_savings, 18, 54
	END IF
	TRANSMIT 'enters the information then a warning message comes up WARNING: ARE YOU SURE YOU WANT TO UPDATE? PF3 TO CANCEL OR TRANSMIT TO UPDATE '
	TRANSMIT 'this confirms the cleared status'
	PF3
	EMReadscreen cleared_confirmation, 1, match_row, 61
	IF cleared_confirmation = "" THEN MsgBox "the match did not appear to clear"
	PF3' this takes us back to DAIL/DAIL

	IF claim_referral_tracking_checkbox = CHECKED Then
	    Call navigate_to_MAXIS_screen ("STAT", "MISC")
	    Row = 6

	    EmReadScreen panel_number, 1, 02, 78
	    If panel_number = "0" then
	    	EMWriteScreen "NN", 20,79
	    	TRANSMIT
	    ELSE
	    	Do
	    		'Checking to see if the MISC panel is empty, if not it will find a new line'
	    		EmReadScreen MISC_description, 25, row, 30
	    		MISC_description = replace(MISC_description, "_", "")
	    		If trim(MISC_description) = "" then
	    			PF9
	    			EXIT DO
	    		Else
	    			row = row + 1
	    		End if
	    	Loop Until row = 17
	    	If row = 17 then MsgBox("There is not a blank field in the MISC panel. Please delete a line(s), and run script again or update manually.")
	    End if
		'writing in the action taken and date to the MISC panel
	 	IF Action_taken_droplist = "CC-Case Closed" or  Action_taken_droplist = "BR-Benefits Reduced"  THEN MISC_action_taken =  "Determination-OP Entered" '"Claim Determination 25 character available
		IF Action_taken_droplist = "NA-No Action Taken" THEN MISC_action_taken = "Determination-No Savings"
		IF Emp_known_droplist = "YES-No Further Action" THEN MISC_action_taken = "Determination-No Savings"
	    EMWriteScreen MISC_action_taken, Row, 30
	    EMWriteScreen date, Row, 66
	    TRANSMIT
	    'PF3
	    start_a_blank_CASE_NOTE
	    Call write_variable_in_case_note("-----Claim Referral Tracking-----")
		Call write_variable_in_case_note("* NDNH new hire information received - " & MISC_action_taken )
	    Call write_bullet_and_variable_in_case_note("Action Date", date)
	    Call write_variable_in_case_note("* Entries for these potential claims must be retained until further notice.")
	    Call write_variable_in_case_note("-----")
	    Call write_variable_in_case_note(worker_signature)
	END IF

	new_hire_first_line = replace(new_hire_first_line, new_HIRE_SSN, "")
	start_a_blank_case_note
	IF Emp_known_droplist = "YES-No Further Action" THEN
		CALL write_variable_in_case_note("-NDNH " & new_hire_first_line & " INFC cleared reported to agency-")
		CALL write_variable_in_case_note("DATE HIRED: " & date_hired)
		CALL write_variable_in_case_note("EMPLOYER: " & employer)
		CALL write_variable_in_case_note(new_hire_third_line)
		CALL write_variable_in_case_note(new_hire_fourth_line)
		CALL write_variable_in_case_note("---")
		CALL write_variable_in_case_note("* Reviewed ECF for requested verifications and MAXIS for correctly budgeted income.")
		CALL write_variable_in_case_note("* Cleared match in INFC/HIRE - Previously reported to agency.")
		CALL write_bullet_and_variable_in_case_note("Other notes", other_notes)
		CALL write_variable_in_case_note("---")
		CALL write_variable_in_case_note(worker_signature)
		PF3
		PF3
	ELSEIF Emp_known_droplist = "NO-See Next Question" THEN
		CALL write_variable_in_case_note("-NDNH " & new_hire_first_line & " INFC cleared unreported to agency-")
		CALL write_variable_in_case_note("DATE HIRED: " & date_hired)
		CALL write_variable_in_case_note("EMPLOYER: " & employer)
		CALL write_variable_in_case_note(new_hire_third_line)
		CALL write_variable_in_case_note(new_hire_fourth_line)
		CALL write_variable_in_case_note("---")
		CALL write_variable_in_case_note("* Reviewed ECF for requested verifications updated INFC/HIRE accordingly")
		IF Action_taken_droplist = "NA-No Action Taken" THEN CALL write_variable_in_case_note("* No futher action taken on this match at this time")
		IF Action_taken_droplist = "BR-Benefits Reduced" THEN CALL write_variable_in_case_note("* Action taken: Benefits Reduced")
		IF Action_taken_droplist = "CC-Case Closed" THEN CALL write_variable_in_case_note("* Action taken: Case Closed (allowing for 10 day cutoff if applicable)")
		IF cost_savings <> "" THEN CALL write_variable_in_case_note("* First Month Cost Savings: $" & cost_savings)
		CALL write_bullet_and_variable_in_case_note("Other notes", other_notes)
		CALL write_variable_in_case_note("---")
		CALL write_variable_in_case_note(worker_signature)
		PF3
		PF3
		IF tenday_checkbox = CHECKED THEN CALL write_variable_in_TIKL("Unable to close due to 10 day cutoff. Verification of job via NEW HIRE should have returned by now. If not received and processed, take appropriate action.")
	END IF
	script_end_procedure_with_error_report("Success! The NDNH HIRE message has been cleared. Please start overpayment process if necessary.")
END IF
 'exampl of the cleared new hire 2279837
TO DO enhancment to make employer name matches
check box for verbal authorization to run work numbers
work number ID add to employer addr
