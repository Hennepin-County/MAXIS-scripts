'Required for statistical purposes==========================================================================================
name_of_script = "NOTICES - NOMI.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 276                     'manual run time in seconds
STATS_denomination = "C"       			   'C is for each CASE
'END OF stats block=========================================================================================================

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

'logic to autofill the 'last_day_for_recert' field
next_month = DateAdd("M", 1, date)
next_month = DatePart("M", next_month) & "/01/" & DatePart("YYYY", next_month)
last_day_for_recert = dateadd("d", -1, next_month) & "" 	'blank space added to make 'last_day_for_recert' a string

'Checks for county info from global variables, or asks if it is not already defined.
get_county_code

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog SNAP_ER_NOMI_dialog, 0, 0, 286, 120, "SNAP ER NOMI dialog"
  EditBox 85, 5, 55, 15, MAXIS_case_number
  EditBox 85, 25, 55, 15, interview_date
  EditBox 225, 25, 55, 15, interview_time
  EditBox 100, 45, 55, 15, last_day_for_recert
  EditBox 100, 70, 180, 15, contact_attempts
  EditBox 70, 95, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 175, 95, 50, 15
    CancelButton 230, 95, 50, 15
  Text 5, 75, 85, 10, "Attempts to contact client:"
  Text 35, 10, 45, 10, "Case number:"
  Text 160, 50, 115, 10, "(Usually the last day of the month)"
  Text 145, 30, 75, 10, "Missed interview time:"
  Text 5, 50, 95, 10, "Recert must be complete by:"
  Text 10, 30, 75, 10, "Missed interview date:"
  Text 5, 100, 60, 10, "Worker signature:"
EndDialog

BeginDialog NOMI_dialog, 0, 0, 261, 125, "NOMI Dialog"
  EditBox 55, 5, 55, 15, MAXIS_case_number
  EditBox 200, 5, 55, 15, application_date
  EditBox 95, 25, 55, 15, interview_date
  EditBox 95, 45, 55, 15, interview_time
  EditBox 95, 65, 160, 15, contact_attempts
  EditBox 70, 85, 75, 15, worker_signature
  CheckBox 10, 110, 205, 10, "Check here to have the script update PND2 for client delay.", client_delay_check
  ButtonGroup ButtonPressed
    OkButton 150, 85, 50, 15
    CancelButton 205, 85, 50, 15
  Text 5, 30, 85, 10, "Date of missed interview:"
  Text 5, 50, 85, 10, "Time of missed interview:"
  Text 140, 10, 55, 10, "Application date:"
  Text 5, 90, 65, 10, "Worker signature:"
  Text 5, 70, 85, 10, "Attempts to contact client:"
  Text 5, 10, 50, 10, "Case number:"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone & grabs case number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

'sets interview time and creates string for variable for Hennepin County recipients
If worker_county_code = "x127" then 
	interview_time = "9:00 AM - 1:00 PM"
	interview_time = interview_time & ""
END IF

'Asks if this is a recert (a recert uses a SPEC/MEMO notice, vs. a SPEC/LETR for intakes and add programs.)
recert_check = MsgBox("Is this a missed SNAP recertification interview?", vbYesNoCancel, "Recertification for SNAP?")
If recert_check = vbCancel then stopscript		'This is the cancel button on a MsgBox
If recert_check = vbYes then 'This is the "yes" button on a MsgBox
	'Sets appointment time for Hennepin County receipients
	DO 
		DO								
			Err_msg = ""
			Dialog SNAP_ER_NOMI_dialog	'dialog for all other users for ER
			cancel_confirmation
			If interview_time = "" then err_msg = err_msg & vbNewLine & "* Select the time of the missed interview."
			If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
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
	Call write_variable_in_SPEC_MEMO("You have missed your Food Support interview that was scheduled for " & interview_date & " at " & interview_time & ".")
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

'If this is not a recert, then APPLICATION verbiage and options are available
Elseif recert_check = vbNo then	'This is the "no" button on a MsgBox
	back_to_self
	'Shows dialog, checks for password prompt
	DO 
		DO					'dialog for all other users
			Err_msg = ""
			Dialog NOMI_dialog
			cancel_confirmation
			If interview_time = "" then err_msg = err_msg & vbNewLine & "* Select the time of the missed interview."
			If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
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
	If client_delay_check = checked then call write_variable_in_CASE_NOTE("* Updated PND2 for client delay.")
	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)
End if

script_end_procedure("Success! The NOMI has been sent, and a case note has been made.")