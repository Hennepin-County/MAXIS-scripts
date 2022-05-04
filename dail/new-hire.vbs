'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.

'Required for statistical purposes===============================================================================
name_of_script = "DAIL - NEW HIRE.vbs"
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
call changelog_update("05/03/2022", "Updated script functionality to support IEVS message updates. This DAIL scrubber will work on both older message with SSN's and new messages without.", "Ilse Ferris, Hennepin County")
call changelog_update("03/01/2020", "Updated TIKL functionality and TIKL text in the case note.", "Ilse Ferris")
call changelog_update("12/17/2019", "Updated navigation to case note from DAIL.", "Ilse Ferris, Hennepin County")
call changelog_update("09/26/2019", "Updated message box regarding children under 19, added policy reference for SNAP/CASH programs.", "Ilse Ferris, Hennepin County")
call changelog_update("08/29/2019", "Added an outlook reminder option.", "MiKayla Handley, Hennepin County")
call changelog_update("08/05/2019", "Updated the term claim referral to use the action taken on MISC and added error reporting.", "MiKayla Handley, Hennepin County")
call changelog_update("08/09/2019", "Added claim referral handling and checks to make sure the case is cleared in INFC.", "MiKayla Handley, Hennepin County")
call changelog_update("01/10/2019", "Updated casenote due to formatting issue, some change to functionality.", "MiKayla Handley, Hennepin County")
call changelog_update("10/17/2018", "Updated the dialog box, no change to functionality.", "MiKayla Handley, Hennepin County")
call changelog_update("05/09/2018", "Updated prospective pay date to the income start date. This resolves an inhibiting error on several cases.", "Ilse Ferris, Hennepin County")
call changelog_update("01/05/2018", "Updated script to ensure TIKL was being created.", "MiKayla Handley, Hennepin County")
call changelog_update("01/05/2018", "Updated coordinates in STAT/JOBS for income type and verification codes.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------------Script
EMConnect ""
still_on_dail = True

'CHECKS TO MAKE SURE THE WORKER IS ON THEIR DAIL
EMReadscreen dail_check, 4, 2, 48
If dail_check <> "DAIL" then script_end_procedure("You are not in your DAIL. This script will stop.")
'TYPES "T" TO BRING THE SELECTED MESSAGE TO THE TOP
EMSendKey "t"
transmit

'determining if the old message with the SSN functionality will be needed or not.
EMReadScreen HIRE_check, 11, 6, 24
If HIRE_check = "JOB DETAILS" then
    SSN_present = True
Else
    EmReadscreen state_match, 4, 6, 20
    If left(state_match, 4) = "SDNH" then SSN_present = False
End if

msgbox "SSN_present: " & SSN_present

'SELECTS THE DAIL MESSAGE AND READS THE RESPONSE
EMSendKey "x"
transmit

EmReadscreen MAXIS_case_number, 8, 6, 57
MAXIS_case_number = Trim(MAXIS_case_number)

'row = 1
'col = 1
'EMSearch "JOB DETAILS", row, col 	'Has to search, because every once in a while the rows and columns can slide one or two positions.

If row = 0 then script_end_procedure_with_error_report("MAXIS may be busy: the script appears to have errored out. This should be temporary. Try again in a moment. If it happens repeatedly contact the alpha user for your agency.")
'Reading information fom the HIRE pop-up
'Date HIred and Employer Name
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
EMReadScreen date_hired, 10, row, col + 12
date_hired = trim(date_hired)
If date_hired = "  -  -  EM" OR date_hired = "UNKNOWN  E" then
    date_hired = CM_mo & "-" & current_day & "-" & CM_yr '??? Why is this code necessary?
Else
    Call ONLY_create_MAXIS_friendly_date(date_hired)
End if
'msgbox date_hired

EMSearch "EMPLOYER:", row, col
EMReadScreen employer, 25, row, col + 10
employer = TRIM(employer)
If SSN_present = True then
    EMReadScreen new_HIRE_SSN, 11, 9, 30
    new_HIRE_SSN = replace(new_HIRE_SSN, "-", "")
Else
    EmReadScreen HH_memb, 2, 9, 35
End if
PF3

If dail_row <> 6 Then Call write_value_and_transmit("t", dail_row, 3)       'bringing the correct message back to the top'

'GOING TO STAT
EMSendKey "s"
transmit
EMReadScreen stat_check, 4, 20, 21
If stat_check <> "STAT" then script_end_procedure_with_error_report("Unable to get to stat due to an error screen. Clear the error screen and return to the DAIL. Then try the script again.")

EMReadScreen stat_case_number, 8, 20, 37
stat_case_number = replace(stat_case_number, "_", "")
If stat_case_number <> MAXIS_case_number Then
	still_on_dail = FALSE
	Call back_to_SELF
	Call navigate_to_MAXIS_screen("STAT", "SUMM")
End If

'GOING TO MEMB, NEED TO CHECK THE HH MEMBER
EMWriteScreen "memb", 20, 71
transmit

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
End if

EMReadScreen memb_age, 2, 8, 76
If cint(memb_age) < 19 then MsgBox "This client is under 19. See CM 0017.15.15 - INCOME OF MINOR CHILD/CAREGIVER UNDER 20 for specific program information about budgeting."

'GOING TO JOBS
EMWriteScreen "jobs", 20, 71
EMWriteScreen HH_memb, 20, 76
transmit

'MFIP cases need to manually add the JOBS panel for ES purposes. '??? why is this here?
If MFIP_case = False then create_JOBS_checkbox = checked
'Setting the variable for the following do...loop
HH_memb_row = 5

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

EMWriteScreen "jobs", 20, 71
EMWriteScreen HH_memb, 20, 76
transmit
msgbox HH_memb

'Checking to see if 5 jobs already exist. If so worker will need to manually delete one first.
EMReadScreen jobs_total_panel_count, 1, 2, 78
IF create_JOBS_checkbox = checked AND jobs_total_panel_count = "5" THEN script_end_procedure_with_error_report("This client has 5 jobs panels already. Please review and delete and unneeded panels if you want the script to add a new one.")

'If new job is known, script ends.
If job_known_checkbox = checked then script_end_procedure("The script will stop as this job is known.")

'Now it will create a new JOBS panel for this case.
If create_JOBS_checkbox = checked then
	EMWriteScreen "nn", 20, 79				'Creates new panel
	transmit

    EmReadscreen closed_case_msg, 27, 20, 79    '??? Not sure if this is how we want to handle these.
    If closed_case_msg = "MAXIS PROGRAMS ARE INACTIVE" then script_end_procedure_with_error_report("This case is inactive. The script will now end.")

	EMReadScreen MAXIS_footer_month, 2, 20, 55	'Reads footer month for updating the panel
	EMReadScreen MAXIS_footer_year, 2, 20, 58		'Reads footer year
	EMWriteScreen "w", 5, 34				'Wage income is the type
	EMWriteScreen "n", 6, 34				'No proof has been provided
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

  	If FS_case = True then 					'If case is SNAP, it creates a PIC
  		EMWriteScreen "x", 19, 38
  		transmit

        IF month_hired = MAXIS_footer_month THEN     'This accounts for rare cases when new hire footer month is the same as the hire date.
            Call create_MAXIS_friendly_date(date_hired, 0, 5, 34) 'Puts date hired if message is from same month as hire ex 01/16 new hire for 1/17/16 start date.
      	ELSE
            Call create_MAXIS_friendly_date(date, 0, 5, 34) 'Puts date hired if message is from same month as hire ex 01/16 new hire for 1/17/16 start date.
      	END IF

        'Entering PIC information
  		EMWriteScreen "1", 5, 64
  		EMWriteScreen "0", 8, 64
  		EMWriteScreen "0", 9, 66
  		transmit
  		transmit
  		transmit
	END IF
	transmit						'Transmits to submit the panel
  	EMReadScreen expired_check, 6, 24, 17 'Checks to see if the jobs panel will carry over by looking for the "This information will expire" at the bottom of the page
  	If expired_check = "EXPIRE" THEN Msgbox "Check next footer month to make sure the JOBS panel carried over"
END IF

If still_on_dail = FALSE Then
	Call back_to_SELF
Else
	PF3 'back to DAIL
End If

'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
Call create_TIKL("Verification of " & employer & "job via NEW HIRE should have returned by now. If not received and processed, take appropriate action. For all federal matches INFC/HIRE must be cleared please see HSR manual.", 10, date, True, TIKL_note_text)

reminder_date = dateadd("d", 10, date)  'Setting out for 10 days reminder
If Outlook_reminder_checkbox = CHECKED THEN
	'Call create_outlook_appointment(appt_date, appt_start_time, appt_end_time, appt_subject, appt_body, appt_location, appt_reminder, appt_category)
	CALL create_outlook_appointment(reminder_date, "08:00 AM", "08:00 AM", "New Hire received " & " for " & MAXIS_case_number, "", "", TRUE, 5, "")
End if

'-----------------------------------------------------------------------------------------CASENOTE
Call navigate_to_MAXIS_screen("CASE", "NOTE")
PF9 ' edit mode
CALL write_variable_in_case_note("--NEW JOB DETAILS FOR (M" & HH_memb & ") unreported to agency-")
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
Call write_variable_in_case_note(TIKL_note_text)
CALL write_bullet_and_variable_in_case_note("Other notes", other_notes)
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)

script_end_procedure_with_error_report("Success! MAXIS updated for new HIRE message, a case note made, and TIKL has been set. An Employment Verification and Verification request form should now be sent. The job is at " & employer & ".")
