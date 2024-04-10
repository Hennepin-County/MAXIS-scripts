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
call changelog_update("04/06/2024", "Removed workaround fix for WAGE match scrubber since the DHS interface with SSN is now repaired.", "Ilse Ferris, Hennepin County")
call changelog_update("06/08/2023", "Fixed bug when entering a job that starting the same month as the HIRE message was generated.", "Ilse Ferris, Hennepin County")
call changelog_update("01/26/2023", "Removed term 'ECF' from the case note per DHS guidance, and referencing the case file instead.", "Ilse Ferris, Hennepin County")
CALL changelog_update("09/16/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
call changelog_update("09/12/2022", "Added support for MEMB 00 messages.", "Ilse Ferris, Hennepin County")
call changelog_update("07/11/2022", "Bug fix in reading the MAXIS Case Number.", "Ilse Ferris, Hennepin County") ''#900
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
EMSendKey "T"   'TYPES "T" TO BRING THE SELECTED MESSAGE TO THE TOP
transmit

'New function that will make sure we are at DAIL, then enter the code to navigate directly from the dail. This can help us keep the tie to the DAIL
function nav_in_DAIL(dail_nav_letter)
    EMReadScreen on_dail_check, 27, 2, 26						'Checking if we are on DAIL - so we can get there
    Do while on_dail_check <> "WORKERS DAILY REPORT (DAIL)"
        PF3														'back up one level
        EMReadScreen SELF_check, 4, 2, 50						'see if we are on SELF
        If SELF_check = "SELF" Then Call navigate_to_MAXIS_screen ("DAIL", "DAIL")	'if we are on SELF - navigate to DAIL/DAIL
        EMReadScreen on_dail_check, 27, 2, 26					'read to see if we made it to DAIL
    Loop

    'To reduce navigation through DAIL, writing case name and number to move quickly
    EMWriteScreen MAXIS_case_number, 20, 38
    EMWriteScreen case_and_person_name, 21, 25
    transmit

    'now we need to find the message that we started with
    dail_row = 6
    Do
        
        'Determining if the script has moved to a new case number within the dail, in which case it needs to move down one more row to get to next dail message
        EMReadScreen new_case, 8, dail_row, 63
        new_case = trim(new_case)
        IF new_case <> "CASE NBR" THEN 
            'If there is NOT a new case number, the script will top the message
            Call write_value_and_transmit("T", dail_row, 3)
        ELSEIF new_case = "CASE NBR" THEN
            'If the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
            Call write_value_and_transmit("T", dail_row + 1, 3)
        End if
        
        'reset the dail_row to 6 so it can navigate through messages in DAIL
        dail_row = 6

        EMReadScreen line_message, 60, dail_row, 20				'read the top message for this case
        line_message = trim(line_message)						'trim the message
        If line_message = full_message Then						'If the message matches the message from the start of the dail scrubber run, then we can enter the info
            Call write_value_and_transmit("X", dail_row, 3)
            EMReadScreen hire_line_1, 60, 9, 5
            EMReadScreen hire_line_2, 60, 10, 5
            EMReadScreen hire_line_3, 60, 11, 5
            EMReadScreen hire_line_4, 60, 12, 5
            full_hire_dail_message_check = hire_line_1 & hire_line_2 & hire_line_3 & hire_line_4
            transmit

            If full_hire_dail_message_check = full_hire_dail_message Then
                EMWriteScreen dail_nav_letter, dail_row, 3
                transmit
                Exit Do
            End If
        End If
        dail_row = dail_row + 1
    Loop
end function

'Resets the DAIL row since the message has now been topped
dail_row = 6  

EmReadScreen MAXIS_case_number, 8, 5, 73
MAXIS_case_number = trim(MAXIS_case_number)

'Read the case name to navigate back again
EMReadScreen case_and_person_name, 8, 5, 5

'Open and read the entire HIRE message to ensure script can find exact match in DAIL again. It then transmits back to same DAIL message
Call write_value_and_transmit("X", 6, 3)
EMReadScreen hire_line_1, 60, 9, 5
EMReadScreen hire_line_2, 60, 10, 5
EMReadScreen hire_line_3, 60, 11, 5
EMReadScreen hire_line_4, 60, 12, 5
full_hire_dail_message = hire_line_1 & hire_line_2 & hire_line_3 & hire_line_4
transmit

If right(full_message, 2) = "00" then
    clear_DAIL = msgbox ("MEMB 00 message are system errors, and only need to be cleared in INFC. Would you like the script to clear this match in INFC?", vbQuestion, vbYesNo, "Non-Actionable DAIL found:")
    If clear_DAIL = vbNo then script_end_procedure("The script will not end. Process the MEMB 00 HIRE message in INFC manually.")
Else
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
End if

'SELECTS THE DAIL MESSAGE AND READS THE RESPONSE
EMSendKey "X"
transmit

'Reading information fom the HIRE pop-up
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
EmReadScreen HH_memb, 2, 9, 15
PF3 ' to exit pop-up

'----------------------------------------------------------------------------------------------------NEW HIRE PORTION
IF match_answer_droplist = "NO-RUN NEW HIRE" THEN
    Call write_value_and_transmit("S", 6, 3)
    
    'PRIV Handling
    EMReadScreen priv_check, 6, 24, 14              'If it can't get into the case then it's a priv case
    If priv_check = "PRIVIL" THEN script_end_procedure("This case is privileged. The script will now end.")
    EMReadScreen stat_check, 4, 20, 21
    
    If stat_check <> "STAT" then script_end_procedure_with_error_report("Unable to get to stat due to an error screen. Clear the error screen and return to the DAIL. Then try the script again.")
    
    'GOING TO MEMB, NEED TO CHECK THE HH MEMBER
    Call write_value_and_transmit("MEMB", 20, 71)
    Call write_value_and_transmit(HH_memb, 20, 76)
    EMReadScreen memb_age, 2, 8, 76
    If cint(memb_age) < 19 then MsgBox "This client is under 19. See CM 0017.15.15 - INCOME OF MINOR CHILD/CAREGIVER UNDER 20 for specific program information about budgeting."

	'GOING TO JOBS
	EMWriteScreen "JOBS", 20, 71
	Call write_value_and_transmit(HH_memb, 20, 76)

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
    		err_msg = ""
    		Dialog Dialog1
    		cancel_confirmation
       		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
    	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

    'Ensure we are at JOBS panel, if not it will navigate back to DAIL to find matching 
    EmReadScreen jobs_panel_check, 17, 2, 33
    If jobs_panel_check <> "Job Income (JOBS)" Then Call nav_in_DAIL("S")

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
            Call create_MAXIS_friendly_date(date_hired, 0, 12, 54) 'Puts date hired if message is from same month as hire ex 01/16 new hire for 1/17/16 start date.
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

    'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
    Call create_TIKL("Verification of " & employer & "job via NEW HIRE should have returned by now. If not received and processed, take appropriate action. For all federal matches INFC/HIRE must be cleared please see HSR manual.", 10, date, True, TIKL_note_text)

    reminder_date = dateadd("d", 10, date)  'Setting out for 10 days reminder
    If Outlook_reminder_checkbox = CHECKED THEN CALL create_outlook_appointment(reminder_date, "08:00 AM", "08:00 AM", "New Hire received for " & MAXIS_case_number, "", "", TRUE, 5, "")

    Call navigate_to_MAXIS_screen("DAIL", "DAIL")
    '-----------------------------------------------------------------------------------------CASE/NOTE
    Call write_value_and_transmit("N", 6, 3)
    PF9
	CALL write_variable_in_case_note("-NDNH Match for (M" & HH_memb & ") for " & trim(employer) & "-")
    CALL write_variable_in_case_note("DATE HIRED: " & date_hired)
    CALL write_variable_in_case_note("EMPLOYER: " & employer)
    CALL write_variable_in_case_note(new_hire_third_line)
    CALL write_variable_in_case_note(new_hire_fourth_line)
    CALL write_variable_in_case_note("---")
    IF ECF_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Sent Verification and Employment verification forms.")
    IF create_JOBS_checkbox = CHECKED THEN CALL write_variable_in_case_note("* STAT/JOBS updated with new hire information from DAIL.")
    IF CCA_checkbox = CHECKED  THEN CALL write_variable_in_case_note("* Sent status update to CCA.")
    IF ES_checkbox = CHECKED  THEN CALL write_variable_in_case_note("* Sent status update to ES.")
    IF work_number_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Sent request for Work Number after confirming client authorization.")
    IF CEI_checkbox = CHECKED THEN CALL write_variable_in_case_note("* Requested CEI/OHI docs.")
    Call write_variable_in_CASE_NOTE(TIKL_note_text)
    CALL write_bullet_and_variable_in_case_note("Other notes", other_notes)
    CALL write_variable_in_case_note("---")
    CALL write_variable_in_case_note(worker_signature)

    script_end_procedure_with_error_report(closing_message)
END IF

'----------------------------------------------------------------------------------------------------INFC Portion
IF match_answer_droplist = "YES-INFC clear match" THEN
    'navigating to the INFC screens
	EMSendKey "I"
	transmit
    'PRIV Handling
    EMReadScreen priv_check, 6, 24, 14              'If it can't get into the case then it's a priv case
    If priv_check = "PRIVIL" THEN script_end_procedure("This case is privileged. The script will now end.")
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

    If clear_DAIL = vbYes then
        'entering the INFC/HIRE match - This is ONLY for MEMB 00 cases
        Call write_value_and_transmit("U", match_row, 3)
        EMReadscreen panel_check, 4, 2, 49
        IF panel_check <> "NHMD" THEN script_end_procedure("The match selected was unable to be entered. The script will now end.")
        EMWriteScreen "N", 16, 54
        EMWriteScreen "NA", 17, 54
        TRANSMIT 'enters the information then a warning message comes up WARNING: ARE YOU SURE YOU WANT TO UPDATE? PF3 TO CANCEL OR TRANSMIT TO UPDATE '
        TRANSMIT 'this confirms the cleared status'
        PF3
        EMReadscreen cleared_confirmation, 1, match_row, 61
        IF cleared_confirmation = " " THEN script_end_procedure("The match selected was unable to be entered. The script will now end.")
        PF3' this takes us back to DAIL/DAIL
    Else
        'This is a dialog asking if the job is known to the agency.
        Dialog1 = ""
        BeginDialog Dialog1, 0, 0, 281, 205, "NDNH Match Resolution Information"
        CheckBox 10, 30, 265, 10, "Check here to verify that ECF has been reviewed and acted upon appropriately", ECF_checkbox
        DropListBox 170, 50, 95, 15, "Select One:"+chr(9)+"YES-No Further Action"+chr(9)+"NO-See Next Question", Emp_known_droplist
        DropListBox 170, 70, 95, 15, "Select One:"+chr(9)+"NA-No Action Taken"+chr(9)+"BR-Benefits Reduced"+chr(9)+"CC-Case Closed", Action_taken_droplist
        EditBox 220, 90, 45, 15, cost_savings
        EditBox 55, 110, 210, 15, other_notes
        CheckBox 10, 140, 260, 10, "Check here if 10 day cutoff has passed - TIKL will be set for following month", tenday_checkbox
        CheckBox 10, 165, 260, 10, "SNAP or MFIP Federal Food only - add Claim Referral Tracking on STAT/MISC", claim_referral_tracking_checkbox
        ButtonGroup ButtonPressed
            OkButton 170, 185, 50, 15
            CancelButton 225, 185, 50, 15
        GroupBox 5, 20, 270, 25, "ECF review"
        Text 10, 55, 145, 10, "Was this employment known to the agency?"
        Text 10, 75, 155, 10, "If unknown: what action was taken by agency?"
        Text 10, 95, 155, 10, "First month cost savings (enter only numbers):"
        Text 10, 115, 40, 10, "Other notes:"
        GroupBox 5, 130, 270, 25, "10 day cutoff for closure"
        GroupBox 5, 155, 270, 25, "Claim Referral Tracking"
        Text 10, 5, 35, 10, "Employer:"
        Text 45, 5, 100, 10, employer
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
	    Call write_value_and_transmit("U", row, 3)
	    EMReadscreen panel_check, 4, 2, 49
	    IF panel_check <> "NHMD" THEN script_end_procedure("The match selected was unable to be entered. The script will now end.")
        EMWriteScreen left(Emp_known_droplist, 1), 16, 54
        If Action_taken_droplist <> "Select One:" then EMWriteScreen left(Action_taken_droplist, 2), 17, 54
	
	    IF cost_savings <> "" THEN
	    	cost_savings = round(cost_savings)
	    	EMWriteScreen cost_savings, 18, 54
	    END IF
	    TRANSMIT 'enters the information then a warning message comes up WARNING: ARE YOU SURE YOU WANT TO UPDATE? PF3 TO CANCEL OR TRANSMIT TO UPDATE '
	    TRANSMIT 'this confirms the cleared status'
	    PF3
	    EMReadscreen cleared_confirmation, 1, match_row, 61
        IF cleared_confirmation = " " THEN script_end_procedure("The match selected was unable to be entered. The script will now end.")
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
	     	IF Action_taken_droplist = "CC-Case Closed" or Action_taken_droplist = "BR-Benefits Reduced"  THEN MISC_action_taken = "Determination-OP Entered" '"Claim Determination 25 character available
	    	IF Action_taken_droplist = "NA-No Action Taken" THEN MISC_action_taken = "Determination-No Savings"
	    	IF Emp_known_droplist = "YES-No Further Action" THEN MISC_action_taken = "Determination-No Savings"
	        EMWriteScreen MISC_action_taken, Row, 30
	        Call write_value_and_transmit(date, Row, 66)

            Call start_a_blank_CASE_NOTE
	        Call write_variable_in_case_note("-----Claim Referral Tracking-----")
	    	Call write_variable_in_case_note("* NDNH new hire information received - " & MISC_action_taken )
	        Call write_bullet_and_variable_in_case_note("Action Date", date)
	        Call write_variable_in_case_note("* Entries for these potential claims must be retained until further notice.")
	        Call write_variable_in_case_note("-----")
	        Call write_variable_in_case_note(worker_signature)
	    END IF

        IF tenday_checkbox = 1 THEN Call create_TIKL("Unable to close due to 10 day cutoff. Verification of job via NEW HIRE should have returned by now. If not received and processed, take appropriate action.", 0, date, True, TIKL_note_text)

        Call start_a_blank_CASE_NOTE
	    IF Emp_known_droplist = "YES-No Further Action" THEN
	    	CALL write_variable_in_case_note("-NDNH Match for (M" & HH_memb & ") INFC cleared: Reported-")
	    	CALL write_variable_in_case_note("DATE HIRED: " & date_hired)
	    	CALL write_variable_in_case_note("EMPLOYER: " & employer)
	    	CALL write_variable_in_case_note(new_hire_third_line)
	    	CALL write_variable_in_case_note(new_hire_fourth_line)
	    	CALL write_variable_in_case_note("---")
	    	CALL write_variable_in_case_note("* Reviewed case file for requested verifications and MAXIS for correctly budgeted income.")
	    	CALL write_variable_in_case_note("* Cleared match in INFC/HIRE - Previously reported to agency.")
	    ELSEIF Emp_known_droplist = "NO-See Next Question" THEN
	    	CALL write_variable_in_case_note("-NDNH Match for (M" & HH_memb & ") INFC cleared: Unreported-")
	    	CALL write_variable_in_case_note("DATE HIRED: " & date_hired)
	    	CALL write_variable_in_case_note("EMPLOYER: " & employer)
	    	CALL write_variable_in_case_note(new_hire_third_line)
	    	CALL write_variable_in_case_note(new_hire_fourth_line)
	    	CALL write_variable_in_case_note("---")
	    	CALL write_variable_in_case_note("* Reviewed case file for requested verifications updated INFC/HIRE accordingly")
	    	IF Action_taken_droplist = "NA-No Action Taken" THEN CALL write_variable_in_case_note("* No further action taken on this match at this time")
	    	IF Action_taken_droplist = "BR-Benefits Reduced" THEN CALL write_variable_in_case_note("* Action taken: Benefits Reduced")
	    	IF Action_taken_droplist = "CC-Case Closed" THEN CALL write_variable_in_case_note("* Action taken: Case Closed (allowing for 10 day cutoff if applicable)")
	    	IF cost_savings <> "" THEN CALL write_variable_in_case_note("* First Month Cost Savings: $" & cost_savings)
        End IF
	    CALL write_bullet_and_variable_in_case_note("Other notes", other_notes)
	    CALL write_variable_in_case_note("---")
	    CALL write_variable_in_case_note(worker_signature)
        closing_message = "Success! The NDNH HIRE message has been cleared. Please start overpayment process if necessary."
	    script_end_procedure_with_error_report(closing_message)
    End if
END IF

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------05/25/2022
'--Tab orders reviewed & confirmed----------------------------------------------05/25/2022
'--Mandatory fields all present & Reviewed--------------------------------------05/25/2022
'--All variables in dialog match mandatory fields-------------------------------05/25/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------05/25/2022-------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------05/25/2022-------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------05/25/2022-------------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------05/25/2022-------------------N/A
'--MAXIS_background_check reviewed (if applicable)------------------------------05/25/2022-------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------05/25/2022
'--Out-of-County handling reviewed----------------------------------------------05/25/2022-------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------05/25/2022
'--BULK - review output of statistics and run time/count (if applicable)--------05/25/2022-------------------N/A
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---05/25/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------05/25/2022
'--Incrementors reviewed (if necessary)-----------------------------------------05/25/2022
'--Denomination reviewed -------------------------------------------------------05/25/2022
'--Script name reviewed---------------------------------------------------------05/25/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------05/25/2022-------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------05/25/2022
'--comment Code-----------------------------------------------------------------05/25/2022
'--Update Changelog for release/update------------------------------------------05/07/2022
'--Remove testing message boxes-------------------------------------------------05/25/2022
'--Remove testing code/unnecessary code-----------------------------------------05/25/2022
'--Review/update SharePoint instructions----------------------------------------05/25/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------05/25/2022
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------05/25/2022
'--Complete misc. documentation (if applicable)---------------------------------05/25/2022
'--Update project team/issue contact (if applicable)----------------------------05/25/2022
