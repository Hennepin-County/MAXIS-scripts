'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "MEMOS - APPOINTMENT LETTER.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
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

'DIALOGS----------------------------------------------------------------------------------------------------
'NOTE: this dialog contains a special modification to allow dynamic creation of the county office list. You cannot edit it in
'	Dialog Editor without modifying the commented line.
BeginDialog appointment_letter_dialog, 0, 0, 156, 355, "Appointment letter"
  EditBox 75, 5, 50, 15, case_number
  DropListBox 50, 25, 95, 15, "new application"+chr(9)+"recertification", app_type
  CheckBox 10, 43, 150, 10, "Check here if this is a reschedule.", reschedule_check
  EditBox 50, 55, 95, 15, CAF_date
  CheckBox 10, 75, 130, 10, "Check here if this is a recert and the", no_CAF_check
  DropListBox 70, 100, 75, 15, "phone"+county_office_list, interview_location		'This line dynamically creates itself based on the information in the FUNCTIONS FILE.
  EditBox 70, 120, 75, 15, interview_date
  EditBox 70, 140, 75, 15, interview_time
  EditBox 80, 160, 65, 15, client_phone
  CheckBox 10, 200, 95, 10, "Client appears expedited", expedited_check
  CheckBox 10, 215, 135, 10, "Same day interview offered/declined", same_day_declined_check
  EditBox 10, 250, 135, 15, expedited_explanation
  CheckBox 10, 280, 135, 10, "Check here if you left V/M with client", voicemail_check
  EditBox 85, 305, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 25, 325, 50, 15
    CancelButton 85, 325, 50, 15
  Text 25, 10, 50, 10, "Case number:"
  Text 15, 30, 30, 10, "App type:"
  Text 15, 60, 35, 10, "CAF date:"
  Text 30, 85, 105, 10, "CAF hasn't been received yet."
  Text 15, 105, 55, 10, "Int'vw location:"
  Text 15, 125, 50, 10, "Interview date: "
  Text 15, 145, 50, 10, "Interview time:"
  Text 15, 160, 60, 20, "Client phone (if phone interview):"
  GroupBox 5, 185, 145, 85, "Expedited questions"
  Text 10, 230, 135, 20, "If expedited interview date is not within six days of the application, explain:"
  Text 45, 290, 75, 10, "requesting a call back."
  Text 15, 310, 65, 10, "Worker signature:"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

'Searches for a case number
call MAXIS_case_number_finder(case_number)


'This Do...loop shows the appointment letter dialog, and contains logic to require most fields.
Do
	Do
		Do
			Do
				Do
					Do
						Do
							Dialog appointment_letter_dialog
							If ButtonPressed = cancel then stopscript
							If isnumeric(case_number) = False or len(case_number) > 8 then MsgBox "You must fill in a valid case number. Please try again."
						Loop until isnumeric(case_number) = True and len(case_number) <= 8 
						CAF_date = replace(CAF_date, ".", "/")
						If no_CAF_check = checked and app_type = "new application" then no_CAF_check = unchecked 'Shuts down "no_CAF_check" so that it will validate the date entered. New applications can't happen if a CAF wasn't provided.
						If no_CAF_check = unchecked and isdate(CAF_date) = False then Msgbox "You did not enter a valid CAF date (MM/DD/YYYY format). Please try again."
					Loop until no_CAF_check = checked or isdate(CAF_date) = True
					if interview_location = "phone" and client_phone = "" then MsgBox "If this is a phone interview, you must enter a phone number! Pleast try again."
				Loop until interview_location <> "phone" or (interview_location = "phone" and client_phone <> "")
				interview_date = replace(interview_date, ".", "/")
				If isdate(interview_date) = False then MsgBox "You did not enter a valid interview date (MM/DD/YYYY format). Please try again."
			Loop until isdate(interview_date) = True 
			If interview_time = "" then MsgBox "You must type an interview time. Please try again."
		Loop until interview_time <> ""
		If no_CAF_check = checked then exit do 'If no CAF was turned in, this layer of validation is unnecessary, so the script will skip it.
		If expedited_check = checked and datediff("d", CAF_date, interview_date) > 6 and expedited_explanation = "" then MsgBox "You have indicated that your case is expedited, but scheduled the interview date outside of the six-day window. An explanation of why is required for QC purposes."
	Loop until expedited_check = unchecked or (datediff("d", CAF_date, interview_date) <= 6) or (datediff("d", CAF_date, interview_date) > 6 and expedited_explanation <> "")
	If worker_signature = "" then MsgBox "You must provide a signature for your case note."
Loop until worker_signature <> ""

'Checks for MAXIS
call check_for_MAXIS(False)

'Using custom function to assign addresses to the selected office
call assign_county_address_variables(county_address_line_01, county_address_line_02)

'Converting the CAF_date variable to a date, for cases where a CAF was turned in
If no_CAF_check = unchecked then CAF_date = cdate(CAF_date)

'Figuring out the last contact day
If app_type = "recertification" then
	next_month = datepart("m", dateadd("m", 1, interview_date))
	next_month_year = datepart("yyyy", dateadd("m", 1, interview_date))
	last_contact_day = dateadd("d", -1, next_month & "/01/" & next_month_year)
End if
If app_type = "new application" then last_contact_day = CAF_date + 30
If DateDiff("d", interview_date, last_contact_day) < 1 then last_contact_day = interview_date 

'This checks to make sure the case is not in background.
MAXIS_background_check

'Navigating to SPEC/MEMO
call navigate_to_MAXIS_screen("SPEC", "MEMO")

'Creates a new MEMO. If it's unable the script will stop.
PF5
EMReadScreen memo_display_check, 12, 2, 33
If memo_display_check = "Memo Display" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")

'Checking for an AREP. If there's an AREP it'll navigate to STAT/AREP, check to see if the forms go to the AREP. If they do, it'll write X's in those fields below.
row = 4								'Defining row and col for the search feature.
col = 1
EMSearch "ALTREP", row, col			'Row and col are variables which change from their above declarations if "ALTREP" string is found.
IF row > 4 THEN						'If it isn't 4, that means it was found.
	arep_row = row											'Logs the row it found the ALTREP string as arep_row
	call navigate_to_MAXIS_screen("STAT", "AREP")			'Navigates to STAT/AREP to check and see if forms go to the AREP
	EMReadscreen forms_to_arep, 1, 10, 45					'Reads for the "Forms to AREP?" Y/N response on the panel.
	call navigate_to_MAXIS_screen("SPEC", "MEMO")			'Navigates back to SPEC/MEMO
	PF5														'PF5s again to initiate the new memo process
END IF
'Checking for SWKR
row = 4								'Defining row and col for the search feature.
col = 1
EMSearch "SOCWKR", row, col 		'Row and col are variables which change from their above declarations if "SOCWKR" string is found.
IF row > 4 THEN						'If it isn't 4, that means it was found.
	swkr_row = row											'Logs the row it found the SOCWKR string as swkr_row
	call navigate_to_MAXIS_screen("STAT", "SWKR")           'Navigates to STAT/SWKR to check and see if forms go to the SWKR
	EMReadscreen forms_to_swkr, 1, 15, 63                   'Reads for the "Forms to SWKR?" Y/N response on the panel.
	call navigate_to_MAXIS_screen("SPEC", "MEMO")           'Navigates back to SPEC/MEMO
	PF5                                                     'PF5s again to initiate the new memo process
END IF
EMWriteScreen "x", 5, 10										'Initiates new memo to client
IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 10		'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 10		'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
transmit														'Transmits to start the memo writing process

'Writes the MEMO.
call write_variable_in_SPEC_MEMO("************************************************************")
IF app_type = "new application" then
	call write_variable_in_SPEC_MEMO("You recently applied for assistance in " & county_name & " on " & CAF_date & ". An interview is required to process your application.")
Elseif app_type = "recertification" then
	If no_CAF_check = unchecked then 
		call write_variable_in_SPEC_MEMO("You sent recertification paperwork to " & county_name & " on " & CAF_date & ". An interview is required to process your application.")
	Else
		call write_variable_in_SPEC_MEMO("You asked us to set up an interview for your recertification. Remember to send in your forms before the interview.")
	End if
End if
call write_variable_in_SPEC_MEMO("")
If interview_location = "phone" then 	'Phone interviews have a different verbiage than any other interview type
	call write_variable_in_SPEC_MEMO("Your phone interview is scheduled for " & interview_date & " at " & interview_time & ".")
Else
	call write_variable_in_SPEC_MEMO("Your in-office interview is scheduled for " & interview_date & " at " & interview_time & ".")
End if
call write_variable_in_SPEC_MEMO("")
If interview_location = "phone" then
	call write_variable_in_SPEC_MEMO("We will be calling you at this number: " & client_phone & ".") 
	call write_variable_in_SPEC_MEMO("")
	call write_variable_in_SPEC_MEMO("If this date and/or time does not work, or you would prefer an interview in the office, please call your worker.")
Else
	call write_variable_in_SPEC_MEMO("Your interview is at the " & interview_location & " Office, located at:")
	call write_variable_in_SPEC_MEMO("   " & county_address_line_01)
	call write_variable_in_SPEC_MEMO("   " & county_address_line_02)
	call write_variable_in_SPEC_MEMO("")
	call write_variable_in_SPEC_MEMO("If this date and/or time does not work, or you would prefer an interview over the phone, please call your worker and provide your phone number.")
End if
call write_variable_in_SPEC_MEMO("")
IF app_type = "new application" then			'"deny your application" vs "your case will auto-close"
	call write_variable_in_SPEC_MEMO("If we do not hear from you by " & last_contact_day & " we will deny your application.")
Elseif app_type = "recertification" then
	call write_variable_in_SPEC_MEMO("If we do not hear from you by " & last_contact_day & ", your case will auto-close.")
END IF
call write_variable_in_SPEC_MEMO("************************************************************")

'Exits the MEMO
PF4

'Navigates to CASE/NOTE and starts a blank one
start_a_blank_CASE_NOTE

'Writes the case note

'If it's rescheduled, that header should engage. Otherwise, it uses separate headers for new apps and recerts.
If reschedule_check = checked then
	call write_variable_in_CASE_NOTE("**Client requested rescheduled appointment, appt letter sent in MEMO.**")
ElseIf app_type = "new application" then 
	call write_variable_in_CASE_NOTE("**New CAF received " & CAF_date & ", appt letter sent in MEMO**")
ElseIf app_type = "recertification" then 
	If no_CAF_check = unchecked then 		'Uses separate headers for whether-or-not a CAF was received.
		call write_variable_in_CASE_NOTE("**Recert CAF received " & CAF_date & ", appt letter sent in MEMO**")
	Else
		call write_variable_in_CASE_NOTE("**Client requested recert appointment, letter sent in MEMO**")	
	End if
End if

'And the rest...
If same_day_declined_check = checked then write_variable_in_CASE_NOTE("* Same day interview offered and declined.")
call write_bullet_and_variable_in_CASE_NOTE("Appointment date", interview_date) 
call write_bullet_and_variable_in_CASE_NOTE("Appointment time", interview_time) 
call write_bullet_and_variable_in_CASE_NOTE("Why interview is more than six days from now", expedited_explanation)
call write_bullet_and_variable_in_CASE_NOTE("Appointment location", interview_location)
call write_bullet_and_variable_in_CASE_NOTE("Client phone", client_phone)
call write_variable_in_CASE_NOTE("* Client must complete interview by " & last_contact_day & ".")
If voicemail_check = checked then call write_variable_in_CASE_NOTE("* Left client a voicemail requesting a call back.")
If forms_to_arep = "Y" then call write_variable_in_CASE_NOTE("* Copy of notice sent to AREP.")				'Defined above
If forms_to_swkr = "Y" then call write_variable_in_CASE_NOTE("* Copy of notice sent to Social Worker.")		'Defined above
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("")
