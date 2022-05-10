'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.As such, it does NOT include protections to be ran independently.

'Required for statistical purposes===============================================================================
name_of_script = "DAIL - NEW HIRE NDNH.vbs"
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
call changelog_update("05/07/2022", "Fixed bug for NDNH new HIRE DAIL's with SSN's in message. There is still a MAXIS system issue that is prohibiting NDNH messages without SSN's to be cleared in INFC. DHS has not provided an update to date.", "Ilse Ferris, Hennepin County")
call changelog_update("05/03/2022", "Updated script functionality to support IEVS message updates. This DAIL scrubber will work on both older message with SSN's and new messages without.", "Ilse Ferris, Hennepin County")
call changelog_update("03/01/2020", "Updated TIKL functionality and TIKL text in the case note.", "Ilse Ferris")
call changelog_update("01/06/2020", "Updated TIKL functionality for TIKL'ing after 10 day cut off.", "Ilse Ferris, Hennepin County")
call changelog_update("12/17/2019", "Updated navigation to case note from DAIL.", "Ilse Ferris, Hennepin County")
call changelog_update("09/26/2019", "Updated message box regarding children under 19, added policy reference for SNAP/CASH programs.", "Ilse Ferris, Hennepin County")
call changelog_update("08/05/2019", "Updated the term claim referral to use the action taken on MISC and added error reporting..", "MiKayla Handley")
call changelog_update("01/10/2019", "Added claim referral handling and checks to make sure the case is cleared in INFC.", "MiKayla Handley, Hennepin County")
call changelog_update("01/10/2019", "Updated casenote due to formatting issue, some change to functionality.", "MiKayla Handley, Hennepin County")
call changelog_update("10/17/2018", "Updated the dialog box, no change to functionality.", "MiKayla Handley, Hennepin County")
CALL changelog_update("06/20/2018", "Updated date handling for hired date and navigation to the DAIL/WRIT.", "MiKayla Handley, Hennepin County")
CALL changelog_update("06/19/2018", "Corrected for handling on duplicate jobs and updated dialog box.", "MiKayla Handley, Hennepin County")
call changelog_update("01/05/2018", "Updated coordinates in STAT/JOBS for income type and verification codes.", "Ilse Ferris, Hennepin County")
call changelog_update("11/30/2017", "Updated NDNH new hire DAIL scrubber with INFC case action handling added.", "MiKayla Handley, Hennepin County")
call changelog_update("09/11/2017", "Initial version.", "MiKayla Handley, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------------Script
EMConnect ""
'CHECKS TO MAKE SURE THE WORKER IS ON THEIR DAIL
EMReadscreen dail_check, 4, 2, 48
If dail_check <> "DAIL" then script_end_procedure("You are not in your DAIL. This script will stop.")
'TYPES "T" TO BRING THE SELECTED MESSAGE TO THE TOP
EMSendKey "T"
transmit

'determining if the old message with the SSN functionality will be needed or not.
EMReadScreen HIRE_check, 11, 6, 37
If HIRE_check = "JOB DETAILS" then
    SSN_present = True
Else
    EmReadscreen fed_match, 4, 6, 20
    If left(fed_match, 4) = "NDNH" then SSN_present = False
    SSN_present = False
End if

'DIALOGS----------------------------------------------------------------------------------------------
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 236, 70, "National Directory of New Hires"
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
		Dialog Dialog1
		Cancel_without_confirmation
		IF match_answer_droplist = "Select One:" THEN err_msg = error_msg & ("You must select an answer.")
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'SELECTS THE DAIL MESSAGE AND READS THE RESPONSE
EMSendKey "X"
transmit

'Reading information fom the HIRE pop-up
EmReadscreen MAXIS_case_number, 8, 6, 57
MAXIS_case_number = Trim(MAXIS_case_number)

'Date Hired and Employer Name
EMReadScreen new_hire_second_line, 61, 10, 5
new_hire_second_line = trim(new_hire_second_line)
'MAXIS Name
EMReadScreen new_hire_third_line, 61, 11, 5
new_hire_third_line = trim(new_hire_third_line)
new_hire_third_line = replace(new_hire_third_line, ",", ", ")
'New Hire Name
EMReadScreen new_hire_fourth_line, 61, 12, 5
new_hire_fourth_line = trim(new_hire_fourth_line)
new_hire_fourth_line = replace(new_hire_fourth_line, ",", ", ")

row = 1 						'Now it's searching for info on the hire date as well as employer
col = 1
'Now it's searching for info on the hire date as well as employer
EMSearch "DATE HIRED", row, col
EMReadScreen date_hired, 10, row, col + 15
date_hired = trim(date_hired)
If date_hired = "  -  -  EM" OR date_hired = "UNKNOWN  E" then
    date_hired = CM_mo & "-" & current_day & "-" & CM_yr '??? Why is this code necessary?
Else
    Call ONLY_create_MAXIS_friendly_date(date_hired)
    month_hired = left(date_hired, 2)       'will be used to determine what dates to use on the JOBS panel
End if

EMSearch "EMPLOYER:", row, col
EMReadScreen employer, 25, row, col + 10
employer = TRIM(employer)
If SSN_present = True then
    EMReadScreen new_HIRE_SSN, 9, 9, 5
Else
    EmReadScreen HH_memb, 2, 9, 15
End if
PF3 ' to exit pop-up

'If the SSN is present and we are just clearing the INFC portion, we don't need to go to STAT. This applies to old NDNH messages that we clearing in INFC.
go_to_STAT = True
If SSN_present = True then
    If match_answer_droplist = "YES-INFC clear match" then
        go_to_STAT = False
    End if
End if

'----------------------------------------------------------------------------------------------------STAT Information
If go_to_STAT = True then
    Call write_value_and_transmit("S", 6, 3)
    EMReadScreen stat_check, 4, 20, 21
    If stat_check <> "STAT" then script_end_procedure_with_error_report("Unable to get to stat due to an error screen. Clear the error screen and return to the DAIL. Then try the script again.")
    'GOING TO MEMB, NEED TO CHECK THE HH MEMBER
    Call write_value_and_transmit("MEMB", 20, 71)

    If SSN_present = True then
        Do
        	EMReadScreen MEMB_current, 1, 2, 73
        	EMReadScreen MEMB_total, 1, 2, 78
        	EMReadScreen MEMB_SSN, 11, 7, 42
        	If new_HIRE_SSN = replace(MEMB_SSN, " ", "") then
                exit do
            Else
        		transmit
            End if
        LOOP UNTIL (MEMB_current = MEMB_total) or (new_HIRE_SSN = replace(MEMB_SSN, " ", ""))
        EMReadScreen HH_memb, 2, 4, 33
    Else
        Call write_value_and_transmit(HH_memb, 20, 76) 'SSN_present = False information here
        EmReadscreen MEMB_SSN, 11, 7, 42    'gathering the SSN
        MEMB_SSN = replace(MEMB_SSN, " ", "")
        If match_answer_droplist = "YES-INFC clear match" then PF3 'back to DAIL
    End if
End if

'----------------------------------------------------------------------------------------------------NEW HIRE PORTION
IF match_answer_droplist = "NO-RUN NEW HIRE" THEN
    EMReadScreen memb_age, 2, 8, 76
    If cint(memb_age) < 19 then MsgBox "This client is under 19. See CM 0017.15.15 - INCOME OF MINOR CHILD/CAREGIVER UNDER 20 for specific program information about budgeting."

	'GOING TO JOBS
	EMWriteScreen "JOBS", 20, 71
	EMWriteScreen HH_memb, 20, 76
	transmit

	create_JOBS_checkbox = checked 'defaulting to checked

    'This is a dialog asking if the job is known to the agency.
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 281, 195, "NEW HIRE INFORMATION"
      EditBox 65, 5, 20, 15, HH_memb
      EditBox 65, 25, 95, 15, employer
      CheckBox 15, 45, 190, 10, "Check here to have the script create a new JOBS panel.", create_JOBS_checkbox
      CheckBox 15, 55, 160, 10, "Job is known to the agency (exit the script).", job_known_checkbox
      CheckBox 15, 80, 195, 10, "Sent a request for verifications out of ECF.", ECF_checkbox
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

'Show dialog
	DO
	    DO
	    	Dialog dialog1
	    	cancel_confirmation
	    	MAXIS_dialog_navigation
	    LOOP UNTIL ButtonPressed = -1
	    call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
    LOOP UNTIL are_we_passworded_out = false

    EMWriteScreen "JOBS", 20, 71    'Ensuring we're on JOBS for the right member still post dialog
	Call write_value_and_transmit(HH_memb, 20, 76)

    'Checking to see if 5 jobs already exist. If so worker will need to manually delete one first.
	EMReadScreen jobs_total_panel_count, 1, 2, 78
	IF create_JOBS_checkbox = checked AND jobs_total_panel_count = "5" THEN script_end_procedure_with_error_report("This client has 5 jobs panels already. Please review and delete and unneeded panels if you want the script to add a new one.")

    'If new job is known, script ends.
	If job_known_checkbox = checked then script_end_procedure("The script will stop as this job is known.")    '??? Should be case noting if the job is known?

    'Now it will create a new JOBS panel for this case.
	If create_JOBS_checkbox = checked then
    	Call write_value_and_transmit("NN", 20, 79)				'Creates new panel

        EmReadscreen closed_case_msg, 27, 20, 79    '??? Not sure if this is how we want to handle these.
        If closed_case_msg = "MAXIS PROGRAMS ARE INACTIVE" then script_end_procedure_with_error_report("This case is inactive. The script will now end.")

    	EMReadScreen MAXIS_footer_month, 2, 20, 55	'Reads footer month for updating the panel
    	EMReadScreen MAXIS_footer_year, 2, 20, 58		'Reads footer year
    	EMWriteScreen "W", 5, 34				'Wage income is the type
    	EMWriteScreen "N", 6, 34				'No proof has been provided
    	EMWriteScreen employer, 7, 42			'Adds employer info

        Call create_MAXIS_friendly_date(date_hired, 0, 9, 35)

      	IF month_hired = MAXIS_footer_month THEN     'This accounts for rare cases when new hire footer month is the same as the hire date.
            Call create_MAXIS_friendly_date(date_hired, 0, 12, 57) 'Puts date hired if message is from same month as hire ex 01/16 new hire for 1/17/16 start date.
      	ELSE
            EmWriteScreen MAXIS_footer_month, 12, 54
      		EMWriteScreen "01", 12, 57		'Puts the first in as the day on prospective side
            EmWriteScreen MAXIS_footer_year, 12, 60
      	END IF

      	EMWriteScreen "0", 12, 67				'Puts $0 in as the received income amt
      	EMWriteScreen "0", 18, 72				'Puts 0 hours in as the worked hours

      	Call write_value_and_transmit("X", 19, 38)
        IF month_hired = MAXIS_footer_month THEN     'This accounts for rare cases when new hire footer month is the same as the hire date.
            Call create_MAXIS_friendly_date(date_hired, 0, 5, 34) 'Puts date hired if message is from same month as hire ex 01/16 new hire for 1/17/16 start date.
        ELSE
            Call create_MAXIS_friendly_date(date, 0, 5, 34) 'Puts date hired if message is from same month as hire ex 01/16 new hire for 1/17/16 start date.
        END IF

        'Entering PIC information - PIC will update no matter is SNAP is active or not. Following steps for coding from POLI TEMP TE02.05.108 Denying/Closing SNAP for No Income Verif
      	EMWriteScreen "1", 5, 64
      	EMWriteScreen "0", 8, 64
      	EMWriteScreen "0", 9, 66
        transmit
        EmReadScreen PIC_warning, 7, 20, 6
        IF PIC_warning = "WARNING" then transmit 'to clear message
        transmit 'back to JOBS panel
        transmit 'to save JOBS panel
        'Adding additional follow up information to the closing message if the data is not likely to carry over to the next footer month.
        EMReadScreen expired_check, 6, 24, 17 'Checks to see if the jobs panel will carry over by looking for the "This information will expire" at the bottom of the page
        closing_message = "Success! MAXIS updated for new HIRE message, a case note made, and a TIKL has been sent for 10 days from now. An Employment Verification and Verif Req Form should now be sent. The job is at " & employer & "."
        If expired_check = "EXPIRE" THEN closing_message = closing_message & vbcr & vbcr & "Check next footer month to make sure the JOBS panel carried over correctly."
	END IF

    Call back_to_SELF

    'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
    Call create_TIKL("Verification of " & employer & "job via NEW HIRE should have returned by now. If not received and processed, take appropriate action. For all federal matches INFC/HIRE must be cleared please see HSR manual.", 10, date, True, TIKL_note_text)

	'-----------------------------------------------------------------------------------------CASENOTE
    Call navigate_to_MAXIS_screen("CASE", "NOTE")
    PF9 ' edit mode
	CALL write_variable_in_case_note("-NDNH JOB DETAILS FOR (M" & HH_memb & ") unreported to agency-")
    CALL write_variable_in_case_note("DATE HIRED: " & date_hired)
    CALL write_variable_in_case_note("EMPLOYER: " & employer)
    CALL write_variable_in_case_note(new_hire_third_line)
    CALL write_variable_in_case_note(new_hire_fourth_line)
    CALL write_variable_in_case_note("---")
    IF ECF_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Sent employment verification and DHS-2919 (Verif Request Form B) from ECF.")
    IF create_JOBS_checkbox = CHECKED THEN CALL write_variable_in_case_note("* STAT/JOBS updated with new hire information from DAIL.")
    IF CCA_checkbox = CHECKED  THEN CALL write_variable_in_case_note("* Sent status update to CCA.")
    IF ES_checkbox = CHECKED  THEN CALL write_variable_in_case_note("* Sent status update to ES.")
    IF work_number_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Sent request for Work Number after confirming client authorization.")
    IF CEI_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Requested CEI/OHI docs.")
    Call write_variable_in_CASE_NOTE(TIKL_note_text)
    CALL write_bullet_and_variable_in_case_note("Other notes", other_notes)
    CALL write_variable_in_case_note("---")
    CALL write_variable_in_case_note(worker_signature)
    PF3

    reminder_date = dateadd("d", 10, date)  'Setting out for 10 days reminder
    If Outlook_reminder_checkbox = CHECKED THEN CALL create_outlook_appointment(reminder_date, "08:00 AM", "08:00 AM", "New Hire received for " & MAXIS_case_number, "", "", TRUE, 5, "")

    script_end_procedure_with_error_report(closing_message)
END IF

'----------------------------------------------------------------------------------------------------INFC Portion
IF match_answer_droplist = "YES-INFC clear match" THEN
    'navigating to the INFC screens
	EMSendKey "I"
	transmit
    If SSN_present = False then EmWriteScreen MEMB_SSN, 3, 63
    Call write_value_and_transmit("HIRE", 20, 71)

    'checking for IRS non-disclosure agreement.
    EMReadScreen agreement_check, 9, 2, 24
    IF agreement_check = "Automated" THEN script_end_procedure("To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the screens and review the agreement.")

	row = 9
	DO
        EMReadScreen case_number, 8, row, 5
        case_number = trim(case_number)
        IF case_number = MAXIS_case_number THEN
    		EMReadScreen employer_match, 20, row, 36
    		employer_match = trim(employer_match)
    		IF trim(employer_match) = "" THEN script_end_procedure("An employer match could not be found. The script will now end.")
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
	IF hire_match <> TRUE THEN script_end_procedure("No pending HIRE match found for: " & employer & "." & vbcr & "Please review case for potential manual updates.")

    'This is a dialog asking if the job is known to the agency.
        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 281, 190, "NDNH Match Resolution Information"
          CheckBox 10, 15, 265, 10, "Check here to verify that ECF has been reviewed and acted upon appropriately", ECF_checkbox
          DropListBox 170, 35, 95, 15, "Select One:"+chr(9)+"YES-No Further Action"+chr(9)+"NO-See Next Question", Emp_known_droplist
          DropListBox 170, 55, 95, 15, "Select One:"+chr(9)+"NA-No Action Taken"+chr(9)+"BR-Benefits Reduced"+chr(9)+"CC-Case Closed", Action_taken_droplist
          EditBox 220, 75, 45, 15, cost_savings
          EditBox 55, 95, 210, 15, other_notes
          CheckBox 10, 125, 260, 10, "Check here if 10 day cutoff has passed - TIKL will be set for following month", tenday_checkbox
		  CheckBox 10, 150, 260, 10, "SNAP or MFIP Federal Food only - add Claim Referral Tracking on STAT/MISC", claim_referral_tracking_checkbox
          ButtonGroup ButtonPressed
            OkButton 170, 170, 50, 15
            CancelButton 225, 170, 50, 15
          GroupBox 5, 5, 270, 25, "ECF review"
          Text 10, 40, 145, 10, "Was this employment known to the agency?"
          Text 10, 60, 155, 10, "If unknown: what action was taken by agency?"
          Text 10, 80, 155, 10, "First month cost savings (enter only numbers):"
          Text 10, 100, 40, 10, "Other notes:"
          GroupBox 5, 115, 270, 25, "10 day cutoff for closure"
          GroupBox 5, 140, 270, 25, "Claim Referral Tracking"
        EndDialog
	DO
		DO
			err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
			Dialog Dialog1
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
	transmit
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
        Call navigate_to_MAXIS_screen("CASE", "NOTE")
        PF9 ' edit mode
	    Call write_variable_in_case_note("-----Claim Referral Tracking-----")
		Call write_variable_in_case_note("* NDNH new hire information received - " & MISC_action_taken )
	    Call write_bullet_and_variable_in_case_note("Action Date", date)
	    Call write_variable_in_case_note("* Entries for these potential claims must be retained until further notice.")
	    Call write_variable_in_case_note("-----")
	    Call write_variable_in_case_note(worker_signature)
	END IF

    Call navigate_to_MAXIS_screen("CASE", "NOTE")
    PF9 ' edit mode
	IF Emp_known_droplist = "YES-No Further Action" THEN
		CALL write_variable_in_case_note("-NDNH JOB DETAILS FOR (M" & HH_memb & ") INFC cleared reported to agency-")
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

	ELSEIF Emp_known_droplist = "NO-See Next Question" THEN
		CALL write_variable_in_case_note("-NDNH JOB DETAILS FOR (M" & HH_memb & ") INFC cleared unreported to agency-")
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
	END IF
    IF tenday_checkbox = 1 THEN Call create_TIKL("Unable to close due to 10 day cutoff. Verification of job via NEW HIRE should have returned by now. If not received and processed, take appropriate action.", 0, date, True, TIKL_note_text)
	script_end_procedure_with_error_report("Success! The NDNH HIRE message has been cleared. Please start overpayment process if necessary.")
END IF
