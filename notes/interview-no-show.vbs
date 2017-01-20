'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - INTERVIEW NO SHOW.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 300          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

BeginDialog same_day_dialog, 0, 0, 191, 278, "Enter No Show Information"
  EditBox 80, 20, 95, 15, MAXIS_case_number
  EditBox 70, 55, 90, 15, interview_date
  EditBox 70, 70, 90, 15, first_page
  EditBox 70, 90, 90, 15, second_page
  CheckBox 15, 120, 155, 20, "Attempted phone call to client - No Answer", pc_attempted
  EditBox 75, 140, 95, 15, time_called
  EditBox 75, 160, 95, 15, phone_number
  CheckBox 75, 175, 90, 15, "Left Message for Client", left_vm
  CheckBox 15, 195, 70, 15, "Potential XFS", potential_xfs
  CheckBox 15, 215, 150, 15, "Check here to have the script send a NOMI", nomi_sent
  EditBox 75, 235, 105, 15, worker_signature
  ButtonGroup ButtonPressed
    CancelButton 35, 255, 70, 15
    OkButton 110, 255, 70, 15
  Text 10, 5, 175, 10, "Client did not respond to page for sameday interview"
  Text 30, 20, 45, 10, "Case Number"
  GroupBox 0, 40, 180, 70, "Client was Paged in the Lobby"
  Text 15, 55, 50, 10, "Interview Date:"
  Text 25, 75, 40, 10, "1st Page at:"
  Text 25, 90, 45, 15, "2nd Page at:"
  GroupBox 0, 110, 180, 85, "Phone Call to Client"
  Text 35, 145, 35, 10, "Called at:"
  Text 15, 160, 50, 15, "Phone Number"
  Text 10, 235, 60, 10, "Worker Signature"
EndDialog

BeginDialog Scheduled_interview_dialog, 0, 0, 231, 280, "Scheduled_interview_dialog"
  EditBox 65, 5, 60, 15, MAXIS_case_number
  EditBox 65, 25, 60, 15, application_date
  DropListBox 65, 45, 165, 15, "Select one..."+chr(9)+"Recertification"+chr(9)+"New Application"+chr(9)+"Recert and Add a program", Type_of_interview_droplist
  CheckBox 5, 85, 30, 10, "Cash", Cash_pend
  CheckBox 45, 85, 30, 10, "SNAP", Fs_pend
  CheckBox 85, 85, 50, 10, "Emergency", Emer_pend
  CheckBox 145, 85, 25, 10, "HC", hc_pend
  CheckBox 180, 85, 35, 10, "GRH", grh_pend
  EditBox 60, 105, 50, 15, interview_date
  EditBox 175, 105, 50, 15, Interview_time
  CheckBox 5, 130, 125, 10, "No Show for In Person Interview", In_person_checkbox
  CheckBox 5, 145, 150, 10, "Attempted Phone Interview- No answer", Phone_int_checkbox
  EditBox 75, 160, 75, 15, Phone_number_scheduled
  CheckBox 35, 180, 120, 10, "Two attempts made to call client", Two_attempts_checkbox
  CheckBox 35, 195, 70, 10, "Left VM for client", Left_VM_checkbox
  CheckBox 5, 240, 55, 10, "Send NOMI ", nomi_sent
  EditBox 35, 215, 190, 15, Case_notes
  EditBox 140, 235, 85, 15, Worker_signature
  ButtonGroup ButtonPressed
    CancelButton 115, 260, 50, 15
    OkButton 175, 260, 50, 15
  Text 5, 10, 45, 10, "Case Number"
  Text 5, 30, 55, 10, "Application Date: "
  Text 5, 50, 60, 10, "Type of interview "
  Text 5, 70, 75, 10, "Programs Applied for: "
  Text 5, 110, 50, 10, "Interview Date"
  Text 120, 110, 50, 10, "Interview Time"
  Text 20, 165, 50, 10, "Phone Number"
  Text 75, 240, 60, 10, "Worker Signature"
  Text 5, 220, 25, 10, "Notes:"
  GroupBox 0, 235, 105, 0, "-5"
  GroupBox -10, -5, 245, 105, ""
  GroupBox 0, 95, 235, 115, ""
EndDialog



BeginDialog SNAP_ER_NOMI_dialog, 0, 0, 211, 102, "SNAP ER NOMI Dialog"
  Text 5, 5, 50, 10, "Case number:"
  EditBox 60, 0, 65, 15, MAXIS_case_number
  Text 5, 25, 85, 10, "Date of missed interview:"
  EditBox 95, 20, 50, 15, Interview_date
  Text 5, 45, 85, 10, "Time of missed interview:"
  EditBox 95, 40, 50, 15, Interview_time
  Text 5, 60, 125, 20, "Recert must be complete by (usually the last day of the current month):"
  EditBox 130, 60, 75, 15, last_day_for_recert
  Text 55, 85, 70, 10, "Sign your case note:"
  EditBox 130, 80, 75, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 155, 5, 50, 15
    CancelButton 155, 25, 50, 15
EndDialog

BeginDialog NOMI_dialog, 0, 0, 151, 155, "NOMI Dialog"
  EditBox 70, 5, 65, 15, MAXIS_case_number
  EditBox 95, 25, 50, 15, Interview_date
  EditBox 95, 45, 50, 15, Interview_time
  EditBox 80, 65, 50, 15, application_date
  EditBox 70, 85, 75, 15, worker_signature
  CheckBox 10, 110, 135, 10, "Check here to have the script update", client_delay_check
  ButtonGroup ButtonPressed
    OkButton 20, 135, 50, 15
    CancelButton 80, 135, 50, 15
  Text 20, 10, 50, 10, "Case number:"
  Text 5, 30, 85, 10, "Date of missed interview:"
  Text 5, 50, 85, 10, "Time of missed interview:"
  Text 20, 70, 60, 10, "Application date:"
  Text 5, 90, 65, 10, "Worker signature:"
  Text 45, 120, 75, 10, "PND2 for client delay."
EndDialog



'THE SCRIPT--------------------------------------------------------------------------------------------------
'Asks if this is a same day interview or not. Scheduled interviews have a different dialog.  The if...then logic will be put in the do...loop.
Same_day_Interview = MsgBox("Is this a same day interview?", vbYesNoCancel or VbDefaultButton2)
If Same_day_Interview = vbCancel then stopscript

'Connects to BlueZone default screen
EMConnect ""
EMFocus

'Pulls case number from MAXIS if worker has already selected a case
Call MAXIS_case_number_finder(MAXIS_case_number)

'Defaults the Interview Date to today's date
interview_date = date & ""

'Defaults the Client Phone number to the first phone number listed on MAXIS in STAT/ADDR
Call navigate_to_MAXIS_screen ("STAT", "ADDR")
EMReadScreen phone_01, 3, 17, 45
EMReadScreen phone_02, 3, 17, 51
EMReadScreen phone_03, 4, 17, 55
phone_number = phone_01 & "-" & phone_02 & "-" & phone_03 & ""


'Determines which programs are currently pending in the month of application
call navigate_to_MAXIS_screen("STAT","PROG")
EMReadScreen cash1_pend, 4, 6, 74
EMReadScreen cash2_pend, 4, 7, 74
EMReadScreen emer_pend, 4, 8, 74
EMReadScreen grh_pend, 4, 9, 74
EMReadScreen fs_pend, 4, 10, 74
EMReadScreen hc_pend, 4, 12, 74

'Assigns a value so the programs pending will show up in check boxes
IF cash1_pend = "PEND" THEN
	cash1_pend = 1
Else
	cash1_pend = 0
End If

If cash2_pend = "PEND" THEN
	cash2_pend = 1
Else
	cash2_pend = 0
End if

If cash1_pend = 1 OR cash2_pend = 1 then cash_pend = 1

If emer_pend = "PEND" THEN
	emer_pend = 1
Else
	emer_pend = 0
End if

If grh_pend = "PEND" THEN
	grh_pend = 1
Else
	grh_pend = 0
End if

If fs_pend = "PEND" THEN
	fs_pend = 1
Else
	fs_pend = 0
End if

If hc_pend = "PEND" THEN
	hc_pend = 1
Else
	hc_pend = 0
End if



'Display's the Dialog Box to imput variable information - includes safeguards for mandatory fields
If same_day_interview = vbYes THEN
	Do
		Do
			Do
				err_msg = ""
				Dialog same_day_dialog
				cancel_confirmation
				IF MAXIS_case_number = "" THEN err_msg = err_msg & vbNewLine & "*Please enter a valid case number"
				IF interview_date = "" THEN err_msg = err_msg & vbNewLine & "*Please enter an Interview Date"
				IF IsDate (interview_date) = False THEN err_msg = err_msg & vbNewLine & "*Please enter a valid Interview Date"
				IF first_page = "" THEN err_msg = err_msg & vbNewLine & "*Please enter the time of the 1st page in the lobby"
				IF second_page = "" THEN err_msg = err_msg & vbNewLine & "*Please enter the time of the second page in the lobby - you must page your client at least twice"
				IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "*Please sign your case note"
				If err_msg <> "" Then msgbox "***NOTICE!!!***" & vbNewLine & err_msg & vbNewLine & vbNewLine & "Please resolve for the script to continue"
			Loop until err_msg = ""
			'The following converts the times entered by the user to a standard format
			IF IsNumeric(first_page) = TRUE THEN
				first_page = FormatNumber (first_page, 2)
				first_page = FormatDateTime (first_page, 4)
			End If
			IF IsNumeric(second_page) = TRUE THEN
				second_page = FormatNumber (second_page, 2)
				second_page = FormatDateTime (second_page ,4)
			End If
			first_page = TimeValue(first_page)
			second_page = TimeValue(second_page)
			'This converts the time to military time for any afternnon times
			If first_page < TimeValue("7:00") THEN first_page = DateAdd("h", 12, first_page)
			If second_page < TimeValue("7:00") THEN second_page = DateAdd("h", 12, second_page)
			'This tests to ensure the page times are at least 15 minutes apart
			IF abs(DateDiff("n", first_page, second_page))<15 THEN MsgBox "You must page client at least 15 minutes apart"
		Loop until abs(DateDiff("n", first_page, second_page))>=15 'and MAXIS_case_number <> "" and interview_date <> "" and IsDate(interview_date) = True and first_page <> "" and second_page <> "" and worker_signature <> ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = false

ELSEIF same_day_interview = vbNo THEN 'Begins dialog if client was no show for scheduled interview
	Do
		Do
			err_msg = ""
			Dialog Scheduled_interview_dialog
			cancel_confirmation
			IF MAXIS_case_number = "" OR (MAXIS_case_number <> "" AND IsNumeric(MAXIS_case_number) = False) THEN err_msg = err_msg & vbNewLine & "*Please enter a valid case number"
			If phone_int_checkbox = 0 and In_person_checkbox = 0 then err_msg = err_msg & vbNewLine & "*Please check either Attempted phone interview or No Show for In Person interview"
			If interview_date = "" Then err_msg = err_msg & vbNewLine & "*Please enter an interview date"
			If Interview_time = "" then err_msg = err_msg & vbNewLine & "*Please enter an interview time"
			If application_date = "" then err_msg = err_msg & vbNewLine & "*Please enter the application date"
			If Phone_int_checkbox = checked and phone_number_scheduled = "" Then err_msg = err_msg & vbNewLine & "Enter the phone number you attempted to call"
			If Type_of_interview_droplist= "Select one..." then err_msg = err_msg & vbNewLine & "*Please select the type of interview"
			If err_msg <> "" Then msgbox "***NOTICE!!!***" & vbNewLine & err_msg & vbNewLine & vbNewLine & "Please resolve for the script to continue"
		Loop Until err_msg = ""
		Call check_for_password(are_we_passworded_out)
	Loop until are_we_passworded_out = false
End if

call check_for_MAXIS(False)

'The NOMI dialog box will pop up if the client checks the box on the dialog box
If nomi_sent = 1 then 'Asks if this is a recert. A recert uses a SPEC/MEMO notice, vs. a SPEC/LETR for intakes and add-a-programs.
	recert_check = MsgBox("Is this a missed SNAP recertification interview?" & Chr(13) & Chr(13) & "If yes, the SNAP missed recert interview notice will be sent. " & Chr(13) & "If no, the regular NOMI will be sent.", 3)
	If recert_check = 2 then stopscript		'This is the cancel button on a MsgBox
	If recert_check = 6 then			'This is the "yes" button on a MsgBox
  'Shows dialog, checks for password prompt

		Do
			Do
				err_msg = ""
				Dialog SNAP_ER_NOMI_dialog
				If ButtonPressed = 0 then stopscript
				If MAXIS_case_number = "" then err_msg = err_msg & vbNewLine & "*You did not enter a case number. Please try again."
				If interview_date = "" then err_msg = err_msg & vbNewLine & "*You did not enter a date of missed interview. Please try again."
				If interview_time = "" then err_msg = err_msg & vbNewLine & "*You did not enter a time of missed interview. Please try again."
				If last_day_for_recert = "" then err_msg = err_msg & vbNewLine & "*You did not enter a date the recert must be completed by. Please try again."
				If worker_signature = "" then err_msg = err_msg & vbNewLine & "*You did not sign your case note. Please try again."
				If err_msg <> "" Then msgbox "***NOTICE!!!***" & vbNewLine & err_msg & vbNewLine & vbNewLine & "Please resolve for the script to continue"
			Loop until err_msg = ""
			Call check_for_password(are_we_passworded_out)
		Loop until are_we_passworded_out = false

		'Navigates into SPEC/MEMO
		call navigate_to_MAXIS_screen("SPEC", "MEMO")

		'Checks to make sure we're past the SELF menu
		EMReadScreen still_self, 27, 2, 28
		If still_self = "Select Function Menu (SELF)" then script_end_procedure("Script was not able to get past SELF menu. Is case in background?")

		'Creates a new MEMO. If it's unable the script will stop.
		PF5
		EMReadScreen memo_display_check, 12, 2, 33
		If memo_display_check = "Memo Display" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")
		EMWriteScreen "x", 5, 10
		transmit

		'Writes the info into the MEMO.
		EMSetCursor 3, 15
		EMSendKey "************************************************************"
		EMSendKey "You have missed your Food Support interview that was scheduled for " & interview_date & " at " & interview_time & "." & "<newline>" & "<newline>"
		EMSendKey "Please contact your worker at the telephone number listed below to reschedule the required Food Support interview." & "<newline>" & "<newline>"
		EMSendKey "The Combined Application Form (DHS-5223), the interview by phone or in the office, and the mandatory verifications needed to process your recertification must be completed by " & last_day_for_recert & " or your Food Support case will Auto-Close on this date." & "<newline>"
		EMSendKey "************************************************************"
		PF4

	Elseif recert_check = 7 then
	'This is the "no" button on a MsgBox 'Shows dialog, checks for password prompt
		Do
			Do
				err_msg = ""
				Dialog NOMI_dialog
				If ButtonPressed = 0 then stopscript
				If MAXIS_case_number = "" then err_msg = err_msg & vbNewLine & "*You did not enter a case number. Please try again."
				If isdate(interview_date) = False then err_msg = err_msg & vbNewLine & "*You did not enter a valid interview date."
				If interview_time = "" then err_msg = err_msg & vbNewLine & "*You did not enter interview time."
				If isdate(application_date) = False then err_msg = err_msg & vbNewLine & "*You did not enter a valid application date. Please try again."
				If worker_signature = "" then err_msg = err_msg & vbNewLine & "You did not sign your case note. Please try again."
				If err_msg <> "" Then msgbox "***NOTICE!!!***" & vbNewLine & err_msg & vbNewLine & vbNewLine & "Please resolve for the script to continue"
			Loop until err_msg = ""
			call check_for_password(are_we_passworded_out)
		Loop until are_we_passworded_out = false

		'Navigates into SPEC/LETR
		call navigate_to_MAXIS_screen("SPEC", "LETR")

		'Checks to make sure we're past the SELF menu
		EMReadScreen still_self, 27, 2, 28
		If still_self = "Select Function Menu (SELF)" then script_end_procedure("Unable to get past the SELF screen. Is your case in background?")

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
		PF4

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
	End if
End if

'Formats the page times and time called to standard hh:mm for case note
first_page = FormatDateTime (first_page, 4)
second_page = FormatDateTime (second_page ,4)
IF IsNumeric(time_called) = TRUE THEN
	time_called = FormatNumber (time_called, 2)
	time_called = FormatDateTime (time_called, 4)
End If


'Creates a variable that lists all the programs pending.
If cash_pend = 1 THEN programs_applied_for = programs_applied_for & "Cash, "
If emer_pend = 1 THEN programs_applied_for = programs_applied_for & "Emergency, "
If grh_pend = 1 THEN programs_applied_for = programs_applied_for & "GRH, "
If fs_pend = 1 THEN programs_applied_for = programs_applied_for & "SNAP, "
If hc_pend = 1 THEN programs_applied_for = programs_applied_for & "HC"
programs_applied_for = left(programs_applied_for, len(programs_applied_for) -2)

'Starts a Case Note
Call start_a_blank_case_note

'Writes the case note for Same day interview no show
If Same_day_Interview = vbYes Then
	call write_variable_in_CASE_NOTE("***Attempted to Page Client in Lobby for Interview - No Show***")
	call write_bullet_and_variable_in_CASE_NOTE("Date of application", application_date)
	call write_bullet_and_variable_in_CASE_NOTE("Client was scheduled for interview", interview_date)
	call write_bullet_and_variable_in_CASE_NOTE("Paged client in the lobby to complete interview at", first_page & " & " & second_page)
	IF pc_attempted = 1 THEN call write_bullet_and_variable_in_CASE_NOTE("Attempted to call client, no answer, called at provided number", phone_number & " at " & time_called)
	IF left_vm = 1 THEN call write_variable_in_CASE_NOTE("* Left Voicemail for Client.")
	IF nomi_sent = 1 THEN call write_variable_in_CASE_NOTE("* Sent NOMI to clt through SPEC/LETR.")
	If client_delay_check = 1 then call write_variable_in_Case_note("* Updated PND2 for client delay.")
	IF potential_xfs = 1 THEN call write_variable_in_CASE_NOTE("* Case is Potentially XFS")
	call write_variable_in_CASE_NOTE("---")
	call write_variable_in_CASE_NOTE(worker_signature)
End if
'Writes the case note for Scheduled interview no show
If Same_day_Interview = vbNo then
	If phone_int_checkbox = 1 then call write_variable_in_CASE_NOTE("Attempted Phone Interview- No Answer")
	If In_person_checkbox = 1 then call write_variable_in_CASE_NOTE("No Show for In Person Interview")
	call write_bullet_and_variable_in_CASE_NOTE("Appointment was scheduled for", interview_date & " at " & interview_time)
	If Phone_int_checkbox = 1 then call write_bullet_and_variable_in_CASE_NOTE("Attempted to call client at", phone_number_scheduled)
	If two_attempts_checkbox = 1 then call write_variable_in_CASE_NOTE("* Multiple attempts were made to contact client by phone")
	If Left_VM_checkbox = 1 then call write_variable_in_CASE_NOTE("* Left Voicemail for client to call back")
	call write_bullet_and_variable_in_CASE_NOTE("Application Date", application_date)
	call write_bullet_and_variable_in_CASE_NOTE("Type of Interview", type_of_interview_droplist)
	call write_bullet_and_variable_in_CASE_NOTE("Requesting", programs_applied_for)
	call write_bullet_and_variable_in_CASE_NOTE("Notes", case_notes)
	If nomi_sent = 1 then call write_variable_in_CASE_NOTE("* Nomi was sent to client")
	If client_delay_check = 1 then call write_variable_in_Case_note("* Updated PND2 for client delay.")
	call write_variable_in_CASE_NOTE("---")
	call write_variable_in_CASE_NOTE(worker_signature)
End if

script_end_procedure ("")
