'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "MEMOS - NOMI.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" OR default_directory = "" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
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
BeginDialog SNAP_ER_NOMI_dialog, 0, 0, 211, 102, "SNAP ER NOMI Dialog"
  Text 5, 5, 50, 10, "Case number:"
  EditBox 60, 0, 65, 15, case_number
  Text 5, 25, 85, 10, "Date of missed interview:"
  EditBox 95, 20, 50, 15, date_of_missed_interview
  Text 5, 45, 85, 10, "Time of missed interview:"
  EditBox 95, 40, 50, 15, time_of_missed_interview
  Text 5, 60, 125, 20, "Recert must be complete by (usually the last day of the current month):"
  EditBox 130, 60, 75, 15, last_day_for_recert
  Text 55, 85, 70, 10, "Sign your case note:"
  EditBox 130, 80, 75, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 155, 5, 50, 15
    CancelButton 155, 25, 50, 15
EndDialog

BeginDialog NOMI_dialog, 0, 0, 151, 155, "NOMI Dialog"
  EditBox 70, 5, 65, 15, case_number
  EditBox 95, 25, 50, 15, date_of_missed_interview
  EditBox 95, 45, 50, 15, time_of_missed_interview
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


'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone & grabs case number
EMConnect ""
Call MAXIS_case_number_finder(case_number)

'Asks if this is a recert. A recert uses a SPEC/MEMO notice, vs. a SPEC/LETR for intakes and add-a-programs.
recert_check = MsgBox("Is this a missed SNAP recertification interview?" & Chr(13) & Chr(13) & "If yes, the SNAP missed recert interview notice will be sent. " & Chr(13) & "If no, the regular NOMI will be sent.", 3)
If recert_check = 2 then stopscript		'This is the cancel button on a MsgBox
If recert_check = 6 then			'This is the "yes" button on a MsgBox
	
	
'Shows dialog, checks for password prompt
Do
	Dialog SNAP_ER_NOMI_dialog
	cancel_confirmation
	If case_number = "" then MsgBox "You did not enter a case number. Please try again."
	If date_of_missed_interview = "" then MsgBox "You did not enter a date of missed interview. Please try again."
	If time_of_missed_interview = "" then MsgBox "You did not enter a time of missed interview. Please try again."
	If last_day_for_recert = "" then MsgBox "You did not enter a date the recert must be completed by. Please try again."
	If worker_signature = "" then MsgBox "You did not sign your case note. Please try again."
Loop until case_number <> "" and date_of_missed_interview <> "" and time_of_missed_interview <> "" and last_day_for_recert <> "" and worker_signature <> ""
	
	'checking for an active MAXIS session
	Call check_for_MAXIS(False)
		
	'Navigates into SPEC/MEMO
	call navigate_to_MAXIS_screen("SPEC", "MEMO")
	'Creates a new MEMO. If it's unable the script will stop.
	PF5
	EMReadScreen memo_display_check, 12, 2, 33
	If memo_display_check = "Memo Display" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")
	EMWriteScreen "x", 5, 10
	transmit

	'Writes the info into the MEMO.
	Call write_variable_in_SPEC_MEMO("************************************************************")
	Call write_variable_in_SPEC_MEMO("You have missed your Food Support interview that was scheduled for " & date_of_missed_interview & " at " & time_of_missed_interview & ".")
	Call write_variable_in_SPEC_MEMO(" ")
	Call write_variable_in_SPEC_MEMO("Please contact your worker at the telephone number listed below to reschedule the required Food Support interview.")
	Call write_variable_in_SPEC_MEMO(" ")
	Call write_variable_in_SPEC_MEMO("The Combined Application Form (DHS-5223), the interview by phone or in the office, and the mandatory verifications needed to process your recertification must be completed by " & last_day_for_recert & " or your Food Support case will Auto-Close on this date.")
	Call write_variable_in_SPEC_MEMO("************************************************************")
	PF4

	'Writes the case note
	call start_a_blank_CASE_NOTE
	Call write_variable_in_CASE_NOTE("**Client missed SNAP recertification interview**")
	Call write_variable_in_CASE_NOTE("* Appointment was scheduled for " & date_of_missed_interview & " at " & time_of_missed_interview & ".")
	Call write_variable_in_CASE_NOTE("* A SPEC/MEMO has been sent to the client informing them of missed interview.")
	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)
	MsgBox "Success! A SPEC/MEMO has been sent with the correct language for a missed SNAP recert. A case note has been made."

Elseif recert_check = 7 then		'This is the "no" button on a MsgBox

	'Shows dialog, checks for password prompt

	Do
		Dialog NOMI_dialog
		cancel_confirmation
		If case_number = "" then MsgBox "You did not enter a case number. Please try again."
		If isdate(date_of_missed_interview) = False then MsgBox "You did not enter a valid date of missed interview. Please try again."
		If time_of_missed_interview = "" then MsgBox "You did not enter a time of missed interview. Please try again."
		If isdate(application_date) = False then MsgBox "You did not enter a valid application date. Please try again."
		If worker_signature = "" then MsgBox "You did not sign your case note. Please try again."
	Loop until case_number <> "" and isdate(date_of_missed_interview) = True and time_of_missed_interview <> "" and isdate(application_date) = True and worker_signature <> ""

	'checks for an active MAXIS session
	Call check_for_MAXIS(False)
	
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

	'THE CASE NOTE
	Call start_a_blank_CASE_NOTE
	CALL write_variable_in_CASE_NOTE("**Client missed SNAP interview**")
	CALL write_variable_in_CASE_NOTE("* Appointment was scheduled for " & date_of_missed_interview & " at " & time_of_missed_interview & ".")
	CALL write_variable_in_CASE_NOTE("* A NOMI has been sent via SPEC/LETR informing them of missed interview.")
	If client_delay_check = checked then call write_variable_in_CASE_NOTE("* Updated PND2 for client delay.")
	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)
	MsgBox "Success! The NOMI has been sent and a case note has been made."
End if

script_end_procedure("")