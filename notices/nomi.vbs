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
  EditBox 85, 25, 55, 15, date_of_missed_interview
  EditBox 225, 25, 55, 15, time_of_missed_interview
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
  EditBox 95, 25, 55, 15, date_of_missed_interview
  EditBox 95, 45, 55, 15, time_of_missed_interview
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

'Hennepin County specific dialogs
BeginDialog Hennepin_application_NOMI, 0, 0, 306, 275, "Hennepin County Application SNAP NOMI"
  DropListBox 90, 10, 80, 15, "Select one..."+chr(9)+"Central/NE"+chr(9)+"North"+chr(9)+"Northwest"+chr(9)+"South MPLS"+chr(9)+"S. Suburban"+chr(9)+"West", region_residence
  EditBox 240, 10, 55, 15, MAXIS_case_number
  EditBox 90, 35, 60, 15, date_of_missed_interview
  EditBox 240, 35, 55, 15, time_of_missed_interview
  EditBox 90, 65, 60, 15, application_date
  DropListBox 225, 65, 70, 15, "Select one..."+chr(9)+"First NOMI"+chr(9)+"Second NOMI", NOMI_selection
  EditBox 90, 90, 205, 15, contact_attempts
  EditBox 90, 115, 205, 15, other_info
  CheckBox 10, 140, 150, 10, "Check here to update PND2 for client delay.", client_delay_check
  CheckBox 10, 155, 285, 10, "Check here if HH has applied for MFIP/DWP, and requires a face-to-face interview.", CASH_check
  EditBox 70, 170, 115, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 190, 170, 50, 15
    CancelButton 245, 170, 50, 15
    PushButton 200, 135, 95, 15, "HSR manual NOMI page", HSR_NOMI_button
  Text 5, 95, 85, 10, "Attempts to contact client:"
  Text 185, 15, 45, 10, "Case number:"
  Text 5, 15, 70, 10, "Region of residence: "
  Text 160, 35, 75, 20, "Missed interview time (if applicable):"
  Text 160, 70, 60, 10, "Which NOMI sent:"
  Text 5, 40, 75, 10, "Missed interview date:"
  Text 15, 120, 60, 10, "Other information:"
  GroupBox 5, 195, 290, 70, "Automatic TIKLs "
  Text 15, 210, 270, 20, "If the 'First NOMI' is being sent to the recipient, A TIKL will be made for 11 days from the date sent."
  Text 15, 235, 275, 30, "If the 'Second NOMI' is being sent to the recipient, A TIKL will be made for 30 (if 1st and 2nd NOMI were sent with enough time to allow 10 days prior to day 30), or 60 days from the application date."
  Text 20, 70, 55, 10, "Application date:"
  Text 5, 175, 60, 10, "Worker signature:"
EndDialog

BeginDialog Hennepin_ER_NOMI, 0, 0, 286, 170, "Hennepin County ER SNAP NOMI"
  EditBox 60, 10, 55, 15, MAXIS_case_number
  DropListBox 200, 10, 80, 15, "Select one..."+chr(9)+"Central/NE"+chr(9)+"North"+chr(9)+"Northwest"+chr(9)+"South MPLS"+chr(9)+"S. Suburban"+chr(9)+"West", region_residence
  EditBox 80, 35, 55, 15, date_of_missed_interview
  EditBox 225, 35, 55, 15, time_of_missed_interview
  EditBox 100, 65, 180, 15, contact_attempts
  EditBox 100, 85, 180, 15, other_info
  EditBox 100, 110, 55, 15, last_day_for_recert
  EditBox 70, 140, 100, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 175, 140, 50, 15
    CancelButton 230, 140, 50, 15
    PushButton 175, 110, 105, 15, "HSR manual NOMI page", HSR_NOMI_button
  Text 125, 15, 70, 10, "Region of residence: "
  Text 145, 35, 75, 25, "Missed interview time: (Don't complete if not applicable.)"
  Text 5, 40, 75, 10, "Missed interview date:"
  Text 5, 115, 95, 10, "Recert must be complete by:"
  Text 5, 125, 115, 10, "(Usually the last day of the month)"
  Text 5, 145, 60, 10, "Worker signature:"
  Text 10, 15, 45, 10, "Case number:"
  Text 5, 70, 85, 10, "Attempts to contact client:"
  Text 30, 90, 60, 10, "Other information:"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone & grabs case number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

'Asks if this is a recert (a recert uses a SPEC/MEMO notice, vs. a SPEC/LETR for intakes and add programs.)
recert_check = MsgBox("Is this a missed SNAP recertification interview?", vbYesNoCancel, "Recertification for SNAP?")
If recert_check = vbCancel then stopscript		'This is the cancel button on a MsgBox
If recert_check = vbYes then 'This is the "yes" button on a MsgBox
	DO
		If worker_county_code = "x127" then		'Hennepin specific ER NOMI
			DO
				DO
					Err_msg = ""
					Dialog Hennepin_ER_NOMI
					cancel_confirmation
					'Opening the the HSR manual to the NOMI page
					IF buttonpressed = HSR_NOMI_button then CreateObject("WScript.Shell").Run("https://dept.hennepin.us/hsphd/manuals/hsrm/Pages/NOMI.aspx")
					IF region_residence = "Select one..." then err_msg = err_msg & vbNewLine & "* Select your client's region of residence."
					If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
					If isdate(date_of_missed_interview) = False then err_msg = err_msg & vbNewLine & "* Enter the date of missed interview."
					If isdate(last_day_for_recert) = False then err_msg = err_msg & vbNewLine & "* Enter a date the recert must be completed by."
					If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
					IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
				LOOP until err_msg = ""
			LOOP until ButtonPressed = -1	
		ELSE		
			DO								
				Err_msg = ""
				Dialog SNAP_ER_NOMI_dialog	'dialog for all other users for ER
				cancel_confirmation
				If time_of_missed_interview = "" then err_msg = err_msg & vbNewLine & "* Select the time of the missed interview."
				If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
				If isdate(date_of_missed_interview) = False then err_msg = err_msg & vbNewLine & "* Enter the date of missed interview."
				If isdate(last_day_for_recert) = False then err_msg = err_msg & vbNewLine & "* Enter a date the recert must be completed by."
				If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
				IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
			LOOP until err_msg = ""
		END IF
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

	If worker_county_code = "x127" then
		'writes in the SPEC/MEMO for Hennepin County users
		Call write_variable_in_SPEC_MEMO("************************************************************")
	    IF time_of_missed_interview <> "" THEN
			Call write_variable_in_SPEC_MEMO("You have missed your SNAP interview that was scheduled for " & date_of_missed_interview & " at " & time_of_missed_interview & ".")
		ELSE
			Call write_variable_in_SPEC_MEMO("You have missed your SNAP interview that was scheduled for " & date_of_missed_interview & ".")
	    END IF
		Call write_variable_in_SPEC_MEMO(" ")
	    Call write_variable_in_SPEC_MEMO("Please contact your worker at 612-596-1300 to complete the required SNAP interview.")
		IF region_residence = "Central/NE" Then
			Call write_variable_in_SPEC_MEMO("You may also come to the Human Services building office to complete an interview. The office is located at: 525 Portland Ave. in Minneapolis. Office hours are Monday through Friday from 8 a.m. to 4:30 p.m.")
		ELSEIF region_residence = "North" Then
			Call write_variable_in_SPEC_MEMO("You may also come to the North Minneapolis office to complete an interview. The office is located at: 1001 Plymouth Ave. Office hours are Monday through Friday from 8 a.m. to 4:30 p.m.")
	    ELSEIF region_residence = "Northwest" Then
			Call write_variable_in_SPEC_MEMO("You may also come into the Brooklyn Center to complete an interview. The office is located at: 7051 Brooklyn Blvd. Office hours are Monday through Friday from 8 a.m. to 4:30 p.m.")
		ELSEIF region_residence = "South MPLS" Then
			Call write_variable_in_SPEC_MEMO("You may also come to the Century Plaza office to complete an interview. The office is located at: 330 S. 12th Street in Minneapolis. Office hours are Monday through Friday from 8 a.m. to 4:30 p.m.")
		ELSEIF region_residence = "S. Suburban" Then
			Call write_variable_in_SPEC_MEMO("You may also come into the Bloomington office complete an interview. The office is located at: 9600 Aldrich Ave. Office hours are Monday, Tuesday, Wednesday and Friday from 8 a.m. to 4:30 p.m. and Thursday from 8 a.m. to 6:30 p.m.")
		ElseIF region_residence = "West" Then
			Call write_variable_in_SPEC_MEMO("You may also come into the Hopkins office to complete an interview. The office is located at: 1011 1st Street S. (in the Wells Fargo building). Office hours are Monday through Friday from 8 a.m. to 4:30 p.m.")
		END IF
		Call write_variable_in_SPEC_MEMO(" ")
	    Call write_variable_in_SPEC_MEMO("The Combined Application Form (DHS-5223), the interview by phone or in the office, and the mandatory verifications needed to process your renewal must be completed by " & last_day_for_recert & ", or your SNAP case will Auto-Close on this date.")
		Call write_variable_in_SPEC_MEMO("************************************************************")
	ELSE
		'Writes the info into the MEMO.
		Call write_variable_in_SPEC_MEMO("************************************************************")
		Call write_variable_in_SPEC_MEMO("You have missed your Food Support interview that was scheduled for " & date_of_missed_interview & " at " & time_of_missed_interview & ".")
		Call write_variable_in_SPEC_MEMO(" ")
		Call write_variable_in_SPEC_MEMO("Please contact your worker at the telephone number listed below to reschedule the required Food Support interview.")
		Call write_variable_in_SPEC_MEMO(" ")
		Call write_variable_in_SPEC_MEMO("The Combined Application Form (DHS-5223), the interview by phone or in the office, and the mandatory verifications needed to process your recertification must be completed by " & last_day_for_recert & " or your Food Support case will Auto-Close on this date.")
		Call write_variable_in_SPEC_MEMO("************************************************************")
	END IF
	PF4

	'Writes the case note for the recert NOMI
	call start_a_blank_CASE_NOTE
	Call write_variable_in_CASE_NOTE("**Client missed SNAP recertification interview**")
	If time_of_missed_interview = "" Then
		Call write_variable_in_CASE_NOTE("* Appointment was scheduled for " & date_of_missed_interview & ".")
	ELSE
		Call write_variable_in_CASE_NOTE("* Appointment was scheduled for " & date_of_missed_interview & " at " & time_of_missed_interview & ".")
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
		If worker_county_code = "x127" then		'Hennepin county specific dialog	
			DO
				DO
					Err_msg = ""
					Dialog Hennepin_application_NOMI
					cancel_confirmation
					'Opening the the HSR manual to the NOMI page
					IF buttonpressed = HSR_NOMI_button then CreateObject("WScript.Shell").Run("https://dept.hennepin.us/hsphd/manuals/hsrm/Pages/NOMI.aspx")
					IF region_residence = "Select one..." then err_msg = err_msg & vbNewLine & "* Select your client's region of residence."
					If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
					If isdate(date_of_missed_interview) = False then err_msg = err_msg & vbNewLine & "* Enter the date of missed interview."
					If isdate(last_day_for_recert) = False then err_msg = err_msg & vbNewLine & "* Enter a date the recert must be completed by."
					If isdate(application_date) = False then err_msg = err_msg & vbNewLine & "* You did not enter a valid application date. Please try again."
					If NOMI_selection = "Select one..." then err_msg = err_msg & vbNewLine & "* Please select which NOMI to send."
					If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
					IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
				LOOP until err_msg = ""
			Loop until ButtonPressed = -1 		
		ELSE
			DO					'dialog for all other users
				Err_msg = ""
				Dialog NOMI_dialog
				cancel_confirmation
				If time_of_missed_interview = "" then err_msg = err_msg & vbNewLine & "* Select the time of the missed interview."
				If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
				If isdate(date_of_missed_interview) = False then err_msg = err_msg & vbNewLine & "* Enter the date of missed interview."
				If isdate(last_day_for_recert) = False then err_msg = err_msg & vbNewLine & "* Enter a date the recert must be completed by."
				If isdate(application_date) = False then err_msg = err_msg & vbNewLine & "* You did not enter a valid application date. Please try again."
				If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
				IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
			LOOP until err_msg = ""
		END IF
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
	Loop until are_we_passworded_out = false					'loops until user passwords back in					

	IF worker_county_code = "x127" then  'sends SPEC/MEMO
		Call navigate_to_MAXIS_screen("SPEC", "MEMO")	'Navigating to SPEC/MEMO
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
		
		'writes in the SPEC/MEMO for Hennepin County users
		Call write_variable_in_SPEC_MEMO("*************APPLICATION INTERVIEW REMINDER*************")
		IF time_of_missed_interview <> "" then
			Call write_variable_in_SPEC_MEMO("You recently applied for assistance in Hennepin County on " & (application_date) & " at " & time_of_missed_interview & ".")
		ELSE
			Call write_variable_in_SPEC_MEMO("You recently applied for assistance in Hennepin County on " & (application_date) & ".")
		END IF
		Call write_variable_in_SPEC_MEMO(" ")
		Call write_variable_in_SPEC_MEMO("An interview is required to process your application. You may be eligible for SNAP benefits within 24 hours of your interview.")
		Call write_variable_in_SPEC_MEMO("You must contact your team to complete the interview within the next 10 days. Please call 612-596-1300 if you would like a phone interview.")
		Call write_variable_in_SPEC_MEMO(" ")
		IF region_residence = "Central/NE" Then
			Call write_variable_in_SPEC_MEMO("You may also come to the Human Services building office to complete an interview. The office is located at: 525 Portland Ave. in Minneapolis. Office hours are Monday through Friday from 8 a.m. to 4:30 p.m.")
		ELSEIF region_residence = "North" Then
			Call write_variable_in_SPEC_MEMO("You may also come to the North Minneapolis office to complete an interview. The office is located at: 1001 Plymouth Ave. Office hours are Monday through Friday from 8 a.m. to 4:30 p.m.")
	    ELSEIF region_residence = "Northwest" Then
			Call write_variable_in_SPEC_MEMO("You may also come into the Brooklyn Center to complete an interview. The office is located at: 7051 Brooklyn Blvd. Office hours are Monday through Friday from 8 a.m. to 4:30 p.m.")
		ELSEIF region_residence = "South MPLS" Then
			Call write_variable_in_SPEC_MEMO("You may also come to the Century Plaza office to complete an interview. The office is located at: 330 S. 12th Street in Minneapolis. Office hours are Monday through Friday from 8 a.m. to 4:30 p.m.")
		ELSEIF region_residence = "S. Suburban" Then
			Call write_variable_in_SPEC_MEMO("You may also come into the Bloomington office complete an interview. The office is located at: 9600 Aldrich Ave. Office hours are Monday, Tuesday, Wednesday and Friday from 8 a.m. to 4:30 p.m. and Thursday from 8 a.m. to 6:30 p.m.")
		ElseIF region_residence = "West" Then
			Call write_variable_in_SPEC_MEMO("You may also come into the Hopkins office to complete an interview. The office is located at: 1011 1st Street S. (in the Wells Fargo building). Office hours are Monday through Friday from 8 a.m. to 4:30 p.m.")
		END IF
		If CASH_check = 1 Then 
			Call write_variable_in_SPEC_MEMO(" ")
			Call write_variable_in_SPEC_MEMO("If you are applying for cash, a face-to-face interview is required. If you have been on cash assistance in the last 12 months, you may not need to be interviewed in the office. You may be able to have a phone interview.")
		END IF
		Call write_variable_in_SPEC_MEMO(" ")
		check_date = dateadd("d", 21, application_date)
		If date =< check_date then 
			Call write_variable_in_SPEC_MEMO(" If we do not hear from you by " & (dateadd("d", 31, application_date)) & ", we will deny your application.")
		ELSEif date > check_date then
			Call write_variable_in_SPEC_MEMO(" If we do not hear from you by " & (dateadd("d", 61, application_date)) & ", we will deny your application.")
		END IF 
		Call write_variable_in_SPEC_MEMO(" ")
		Call write_variable_in_SPEC_MEMO("Auth: Laws of Minnesota 7CFR 273.2(e)(3)")
		Call write_variable_in_SPEC_MEMO(" ")
		Call write_variable_in_SPEC_MEMO("If you cannot attend an interview because of a hardship, please call our office.")
		Call write_variable_in_SPEC_MEMO("************************************************************")
		PF4
	ELSE
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
		call create_MAXIS_friendly_date(date_of_missed_interview, 0, 14, 38)
		transmit
		PF4 	'saves the MEMO/LETR
	END IF
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
	
	'First NOMI TIKL
	If NOMI_selection = "First NOMI" then 
		call navigate_to_MAXIS_screen("dail", "writ")
		call create_MAXIS_friendly_date(date, 11, 5, 18) 
		Call write_variable_in_TIKL(NOMI_selection & " was sent 11 days ago. Check case notes to see if interview has been conducted. Send the Second NOMI if client has not completed the interview to date.")
		transmit	
		PF3
	End if
	
	'date variables for the Second NOMI TIKL
	check_date = dateadd("d", 21, application_date)
	pending_sixty_days_date = dateadd("d", 60, application_date)
	ten_day_date = dateadd("d", 10, date)
	
	'Second NOMI TIKL
	If NOMI_selection = "Second NOMI" then  
		call navigate_to_MAXIS_screen("DAIL", "WRIT")
		IF date =< check_date then
			days_pending = "a full 30 days"
			call create_MAXIS_friendly_date(application_date, 31, 5, 18)		'sets a 30 day pending TIKL if the date at least 10 days exist between the NOMI sent and day 30
		ELSEif pending_sixty_days_date =< date then
			call create_MAXIS_friendly_date(date, 10, 5, 18)	'sets a 10 day TIKL if the current date is less than the over
			days_pending = "another 10 days"
		ELSEif pending_sixty_days_date < ten_day_date then
			call create_MAXIS_friendly_date(date, 10, 5, 18)	'sets a 10 day TIKL if the current date is less than the over
			days_pending = "60 days allowing for 10 day notice"
		else 
			call create_MAXIS_friendly_date(application_date, 61, 5, 18)		'otherwise a 60 day TIKL is set
			days_pending = "60 days from the application date"
		END IF
		Call write_variable_in_TIKL(NOMI_selection & " was sent & case has been pending for " & days_pending & ". Check case notes to see if interview has been conducted. Deny the case if the client has not completed the interview.")
		transmit
		PF3
	END IF
 	
	'THE CASE NOTE
	Call start_a_blank_CASE_NOTE
	CALL write_variable_in_CASE_NOTE("**Client missed SNAP interview**")
	If time_of_missed_interview = "" Then
		Call write_variable_in_CASE_NOTE("* Appointment was scheduled for " & date_of_missed_interview & ".")
	ELSE
		Call write_variable_in_CASE_NOTE("* Appointment was scheduled for " & date_of_missed_interview & " at " & time_of_missed_interview & ".")
	END IF
	Call write_bullet_and_variable_in_CASE_NOTE("Attempts to contact the client", contact_attempts)
	Call write_bullet_and_variable_in_CASE_NOTE("Other information", other_info)
	IF worker_county_code = "x127" then
		CALL write_variable_in_CASE_NOTE("* A NOMI has been sent via SPEC/MEMO informing them of missed interview.")
	ELSE
		CALL write_variable_in_CASE_NOTE("* A NOMI has been sent via SPEC/LETR informing them of missed interview.")
	END IF
	If client_delay_check = checked then call write_variable_in_CASE_NOTE("* Updated PND2 for client delay.")
	If NOMI_selection = "First NOMI" then 
		Call write_variable_in_CASE_NOTE("* First NOMI has been sent")
		Call write_variable_in_CASE_NOTE("* A TIKL has been made for 11 days from now to follow-up on application progress.")
	END IF
	If NOMI_selection = "Second NOMI" then 
		Call write_variable_in_CASE_NOTE("* Second NOMI has been sent")
		call write_variable_in_CASE_NOTE("* A TIKL has been made for " & days_pending & " to follow-up on application progress.")
	END IF 
	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)
End if

script_end_procedure("Success! The NOMI has been sent, and a case note has been made.")
