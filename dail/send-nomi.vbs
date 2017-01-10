'Required for statistical purposes===============================================================================
name_of_script = "DAIL - SEND NOMI.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 276         'manual run time in seconds
STATS_denomination = "C"       'C is for case
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
call changelog_update("01/10/2017", "Updated TIKL functionality. A TIKL is created for Application Day 30 if NOMI is sent prior to Application Day 30. Otherwise a TIKL is created for an additional 10 days .", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Resolved merge conflict error.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Updated script to support TIKL created by the NOTES - APPOINTMENT LETTER script. Also removed Hennepin County specific NOMI process. ", "Ilse Ferris, Hennepin County")
call changelog_update("11/20/2016", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Checks for county info from global variables, or asks if it is not already defined.
get_county_code

'logic to autofill the 'last_day_for_recert' into the notice
next_month = DateAdd("M", 1, date)
next_month = DatePart("M", next_month) & "/01/" & DatePart("YYYY", next_month)
last_day_for_recert = dateadd("d", -1, next_month) & "" 	'blank space added to make 'last_day_for_recert' a string

'Dialogs----------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 206, 65, "Case number dialog"
  EditBox 100, 10, 55, 15, MAXIS_case_number
  CheckBox 5, 30, 195, 10, "Check here if missed interview is for SNAP/MFIP renewal.", recert_checkbox
  ButtonGroup ButtonPressed
    OkButton 55, 45, 50, 15
    CancelButton 110, 45, 50, 15
  Text 50, 15, 45, 10, "Case number:"
EndDialog

BeginDialog SNAP_ER_NOMI_dialog, 0, 0, 286, 105, "SNAP ER NOMI dialog"
  EditBox 85, 10, 55, 15, interview_date
  EditBox 225, 10, 55, 15, interview_time
  EditBox 100, 30, 55, 15, last_day_for_recert
  EditBox 100, 55, 180, 15, contact_attempts
  EditBox 70, 80, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 175, 80, 50, 15
    CancelButton 230, 80, 50, 15
  Text 5, 60, 85, 10, "Attempts to contact client:"
  Text 160, 35, 115, 10, "(Usually the last day of the month)"
  Text 145, 15, 75, 10, "Missed interview time:"
  Text 5, 35, 95, 10, "Recert must be complete by:"
  Text 10, 15, 75, 10, "Missed interview date:"
  Text 5, 85, 60, 10, "Worker signature:"
EndDialog

BeginDialog NOMI_dialog, 0, 0, 261, 125, "NOMI Dialog"
  EditBox 95, 5, 70, 15, application_date
  EditBox 95, 25, 70, 15, interview_date
  EditBox 95, 45, 70, 15, interview_time
  EditBox 95, 65, 160, 15, contact_attempts
  EditBox 70, 85, 75, 15, worker_signature
  CheckBox 10, 110, 205, 10, "Check here to have the script update PND2 for client delay.", client_delay_check
  ButtonGroup ButtonPressed
    OkButton 150, 85, 50, 15
    CancelButton 205, 85, 50, 15
  Text 5, 30, 85, 10, "Date of missed interview:"
  Text 5, 50, 85, 10, "Time of missed interview:"
  Text 35, 10, 55, 10, "Application date:"
  Text 5, 90, 65, 10, "Worker signature:"
  Text 5, 70, 85, 10, "Attempts to contact client:"
EndDialog

'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER; As such, it does NOT include protections to be ran independently.
EMConnect ""
EMSendKey "x"
transmit

'Reading date and time of recertification appointment from the TIKL--DAIL message that should be read is: "~*~*~CLIENT HAD RECERT INTERVIEW APPT AT..." This is the part that is static in the DAIL message
EMReadScreen TIKL_content, 16, 9, 17
If TIKL_content = "WAS SENT AN APPT" then
	EMReadScreen interview_info, 19, 10, 5
Else
	EMReadScreen interview_info, 19, 9, 46    'reads "MM/DD/YYYY HH:MM PM" (or any combination less) off of dail messate
End if

'searching for case number
row  = 1
col = 1
EMSearch "Case Number: ", row, col
If row =- 0 then script_end_procedure("MAXIS may be busy: the script appears to have errored out. This should be temporary. Try again in a moment. If it happens repeatedly contact the alpha user for your agency.")
EMReadScreen MAXIS_case_number, 8, row, col + 12
MAXIS_case_number = trim(MAXIS_case_number)
PF3 			'removes the TIKL window
'navigates to CASE/NOTE to user can see if interview has been completed or not
EMSendKey "n"
transmit

'Formatting the interview info to autofill into the NOMI
interview_info = trim(interview_info)                    	'trimming the interview info information
If instr(interview_info, " ") then    					 	'Most cases have both interview date and interview time. This seperates them.
	length = len(interview_info)                           	'establishing the length of the variable
	position = InStr(interview_info, " ")                  	'sets the position at the deliminator (in this case a space)
	interview_date = Left(interview_info, position-1)       'establishes the interview date as being before the deliminator
	If worker_county_code = "x127" then
		interview_time = "9:00 AM - 1:00 PM"				'Hennepin county has custom times for interviews
	Else
		interview_time = Right(interview_info, length-position) 'establishes interview time as after before the deliminator
	END IF
End If

'Msgbox asking the user misssed their interview
interview_confirm = MsgBox("Has an interview been completed for this case?", vbYesNoCancel + vbQuestion, "Interview confirmation")
If interview_confirm = vbCancel then stopscript
If interview_confirm = vbYes then  			'returns user back to DAIL/DAIL and stops the script since no further action is required
	PF3
	script_end_procedure("Success! A NOMI is not required if the interview has been completed." & vbNewLine & "Please review the case for completion if necessary.")
ELSEIF interview_confirm = vbNo then 		'interview was not completed
	Do
    	Do
    		err_msg = ""
    		dialog case_number_dialog
    		If ButtonPressed = 0 then stopscript
    		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
    	Loop until err_msg = ""
    call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
    LOOP UNTIL are_we_passworded_out = false

    'If user selected the renewal checkbox, then the SPEC/MEMO for renewals will be sent.
    If recert_checkbox = 1 then
    	DO
    		DO
    			Err_msg = ""
    			Dialog SNAP_ER_NOMI_dialog	'dialog for all other users for ER
    			cancel_confirmation
    			If worker_county_code <> "x127" and interview_time = "" then err_msg = err_msg & vbNewLine & "* Select the time of the missed interview."
    			If isdate(interview_date) = False then err_msg = err_msg & vbNewLine & "* Enter the date of missed interview."
    			If isdate(last_day_for_recert) = False then err_msg = err_msg & vbNewLine & "* Enter a date the recert must be completed by."
    			If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
    			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    		LOOP until err_msg = ""
    		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    	Loop until are_we_passworded_out = false					'loops until user passwords back in

    	call navigate_to_MAXIS_screen("SPEC", "MEMO")		'Navigating to SPEC/MEMO
    	'Creates a new MEMO. If it's unable the script will stop.
    	PF5
    	EMReadScreen memo_display_check, 12, 2, 33
    	If memo_display_check = "Memo Display" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")

    	'Checking for an AREP. If there's an AREP it'll navigate to STAT/AREP, check to see if the forms go to the AREP. If they do, it'll write X's in those fields below.
    	row = 4                             'Defining row and col for the search feature.
    	col = 1
    	EMSearch "ALTREP", row, col         'Row and col are variables which change from their above declarations if "ALTREP" string is found.
    	IF row > 4 THEN                     'If it isn't 4, that means it was found.
    		arep_row = row                                          'Logs the row it found the ALTREP string as arep_row
    		call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
    		EMReadscreen forms_to_arep, 1, 10, 45                   'Reads for the "Forms to AREP?" Y/N response on the panel.
    		call navigate_to_MAXIS_screen("SPEC", "MEMO")           'Navigates back to SPEC/MEMO
    		PF5                                                     'PF5s again to initiate the new memo process
    	END IF

    	'Checking for SWKR
    	row = 4                             'Defining row and col for the search feature.
    	col = 1
    	EMSearch "SOCWKR", row, col         'Row and col are variables which change from their above declarations if "SOCWKR" string is found.
    	IF row > 4 THEN                     'If it isn't 4, that means it was found.
    		swkr_row = row                                          'Logs the row it found the SOCWKR string as swkr_row
    		call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
    		EMReadscreen forms_to_swkr, 1, 15, 63                'Reads for the "Forms to SWKR?" Y/N response on the panel.
    		call navigate_to_MAXIS_screen("SPEC", "MEMO")         'Navigates back to SPEC/MEMO
    		PF5                                           'PF5s again to initiate the new memo process
    	END IF
    	EMWriteScreen "x", 5, 10                                        'Initiates new memo to client
    	IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
    	IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
    	transmit

    	'Writes the info into the MEMO.
    	Call write_variable_in_SPEC_MEMO("************************************************************")
    	If interview_time = "" Then
    		Call write_variable_in_SPEC_MEMO("You have missed your Food Support interview that was scheduled for " & interview_date & ".")
    	Else
    		Call write_variable_in_SPEC_MEMO("You have missed your Food Support interview that was scheduled for " & interview_date & " at " & interview_time & ".")
    	End if
    	Call write_variable_in_SPEC_MEMO(" ")
    	Call write_variable_in_SPEC_MEMO("Please contact your worker at the telephone number listed below to reschedule the required Food Support interview.")
    	Call write_variable_in_SPEC_MEMO(" ")
    	Call write_variable_in_SPEC_MEMO("The Combined Application Form (DHS-5223), the interview by phone or in the office, and the mandatory verifications needed to process your recertification must be completed by " & last_day_for_recert & " or your Food Support case will Auto-Close on this date.")
    	Call write_variable_in_SPEC_MEMO("************************************************************")
    	PF4

    	'Writes the case note for the recert NOMI
    	call start_a_blank_CASE_NOTE
    	Call write_variable_in_CASE_NOTE("**Client missed SNAP recertification interview**")
    	If interview_time = "" Then
    		Call write_variable_in_CASE_NOTE("* Appointment was scheduled for " & interview_date & ".")
    	ELSE
    		Call write_variable_in_CASE_NOTE("* Appointment was scheduled for " & interview_date & " at " & interview_time & ".")
    	END IF
    	Call write_bullet_and_variable_in_CASE_NOTE("Attempts to contact the client", contact_attempts)
    	Call write_variable_in_CASE_NOTE("* A SPEC/MEMO has been sent to the client informing them of missed interview.")
    	Call write_bullet_and_variable_in_CASE_NOTE("Other information", other_info)
    	Call write_variable_in_CASE_NOTE("* Case will auto-close on " & last_day_for_recert & " if recertification is not completed.")
    	Call write_variable_in_CASE_NOTE("---")
    	Call write_variable_in_CASE_NOTE(worker_signature)

    	PF3	'saves the case note'
    	Call navigate_to_MAXIS_screen("DAIL", "DAIL") 'brings user back to DAIL'
    	script_end_procedure("Success! A SNAP NOMI for recertification SPEC/MEMO has been sent, and a case note has been created.")

    'If this is not a recert, then APPLICATION verbiage and options are available
    Else
    	back_to_self
    	'sets interview time and creates string for variable for Hennepin County recipients
    	If worker_county_code = "x127" then
    		interview_time = "9:00 AM - 1:00 PM"
    		interview_time = interview_time & ""
    	END IF

    	'grabs CAF date, turns CAF date into string for variable
    	call autofill_editbox_from_MAXIS(HH_member_array, "PROG", application_date)
    	application_date = application_date & ""

    	'Shows dialog, checks for password prompt
    	DO
    		DO					'dialog for all other users
    			Err_msg = ""
    			Dialog NOMI_dialog
    			cancel_confirmation
    			If interview_time = "" then err_msg = err_msg & vbNewLine & "* Select the time of the missed interview."
    			If isdate(interview_date) = False then err_msg = err_msg & vbNewLine & "* Enter the date of missed interview."
    			If isdate(last_day_for_recert) = False then err_msg = err_msg & vbNewLine & "* Enter a date the recert must be completed by."
    			If isdate(application_date) = False then err_msg = err_msg & vbNewLine & "* You did not enter a valid application date. Please try again."
    			If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
    			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    		LOOP until err_msg = ""
    		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    	Loop until are_we_passworded_out = false					'loops until user passwords back in

    	'Navigates into SPEC/LETR
    	call navigate_to_MAXIS_screen("SPEC", "LETR")
    	'Opens up the NOMI LETR. If it's unable the script will stop.
    	EMWriteScreen "x", 7, 12
    	transmit
    	EMReadScreen LETR_check, 4, 2, 49
    	If LETR_check = "LETR" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")

    	'Writes the info into the NOMI.
    	EMWriteScreen "x", 7, 17
    	call create_MAXIS_friendly_date(application_date, 0, 12, 38)
    	call create_MAXIS_friendly_date(interview_date, 0, 14, 38)
    	transmit
    	PF4 	'saves the MEMO/LETR

    	'Navigates to REPT/PND2 and updates for client delay if applicable.
    	If client_delay_check = checked then
    		call navigate_to_MAXIS_screen("rept", "pnd2")
    		EMGetCursor PND2_row, PND2_col
    		for i = 0 to 1 'This is put in a for...next statement so that it will check for "additional app" situations, where the case could be on multiple lines in REPT/PND2. It exits after one if it can't find an additional app.
    			EMReadScreen PND2_SNAP_status_check, 1, PND2_row, 62
    			If PND2_SNAP_status_check = "P" then EMWriteScreen "C", PND2_row, 62
    			EMReadScreen PND2_HC_status_check, 1, PND2_row, 6
    			If PND2_HC_status_check = "P" then
    				EMWriteScreen "x", PND2_row, 3
    				transmit
    				person_delay_row = 7
    				Do
    					EMReadScreen person_delay_check, 1, person_delay_row, 39
    					If person_delay_check <> " " then EMWriteScreen "c", person_delay_row, 39
    					person_delay_row = person_delay_row + 2
    				Loop until person_delay_check = " " or person_delay_row > 20
    				PF3
    			End if
    			EMReadScreen additional_app_check, 14, PND2_row + 1, 17
    			If additional_app_check <> "ADDITIONAL APP" then exit for
    			PND2_row = PND2_row + 1
    		next
    		PF3
    		EMReadScreen PND2_check, 4, 2, 52
    		If PND2_check = "PND2" then
    			MsgBox "PND2 might not have been updated for client delay. There may have been a MAXIS error. Check this manually after case noting."
    			PF10
    			client_delay_check = 0
    		End if
    	End if

        'date variables for the TIKL
        day30_date = dateadd("d", 31, application_date)

    	'Sets TIKL
        call navigate_to_MAXIS_screen("DAIL", "WRIT")
        IF date < day30_date then											'if current date is less than the application date 
        	days_pending = "30 days"										'value of variable for case note & TIKL to "30 days"
        	call create_MAXIS_friendly_date(application_date, 31, 5, 18)	'sets a 30 day pending TIKL if the date if at least 10 days exists between the NOMI sent and pending day 30
        ELSE 
			days_pending = "10 additional days"								'value of variable for case note & TIKL to "10 additional days"
			call create_MAXIS_friendly_date(date, 10, 5, 18)				'sets a 10 day TIKL if the current date is equal to over over the application date 
        END IF
		
        Call write_variable_in_TIKL("A NOMI was sent & case has been pending for " & days_pending & ". Check case notes to see if interview has been completed. Deny the case if the client has not completed the interview.")
        transmit
        PF3

    	'THE CASE NOTE
    	Call start_a_blank_CASE_NOTE
    	CALL write_variable_in_CASE_NOTE("**Client missed SNAP interview**")
    	If interview_time = "" Then
    		Call write_variable_in_CASE_NOTE("* Appointment was scheduled for " & interview_date & ".")
    	ELSE
    		Call write_variable_in_CASE_NOTE("* Appointment was scheduled for " & interview_date & " at " & interview_time & ".")
    	END IF
    	Call write_bullet_and_variable_in_CASE_NOTE("Attempts to contact the client", contact_attempts)
    	Call write_bullet_and_variable_in_CASE_NOTE("Other information", other_info)
    	CALL write_variable_in_CASE_NOTE("* A NOMI has been sent via SPEC/LETR informing them of missed interview.")
    	call write_variable_in_CASE_NOTE("* A TIKL has been made for " & days_pending & " to follow-up on the application's progress.")
    	If client_delay_check = checked then call write_variable_in_CASE_NOTE("* Updated PND2 for client delay.")
    	Call write_variable_in_CASE_NOTE("---")
    	Call write_variable_in_CASE_NOTE(worker_signature)

    	PF3	'saves the case note'
    	Call navigate_to_MAXIS_screen("DAIL", "DAIL") 'brings user back to DAIL'
    	script_end_procedure("Success! A NOMI has been sent, and a case note and TIKL have been created.")
    END IF
END IF