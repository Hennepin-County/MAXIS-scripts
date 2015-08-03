'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - LOBBY NO SHOW.vbs"
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

BeginDialog no_show_dialog, 0, 0, 191, 278, "Enter No Show Information"
  EditBox 80, 20, 95, 15, case_number
  EditBox 70, 57, 90, 15, interview_date
  EditBox 70, 73, 90, 15, first_page
  EditBox 70, 90, 90, 15, second_page
  CheckBox 16, 124, 152, 20, "Attempted phone call to client - No Answer", pc_attempted
  EditBox 75, 144, 95, 15, time_called
  EditBox 75, 161, 95, 15, phone_number
  CheckBox 75, 179, 86, 15, "Left Message for Client", left_vm
  CheckBox 16, 198, 70, 15, "Potential XFS", potential_xfs
  CheckBox 16, 215, 150, 15, "Check here to have the script send a NOMI", nomi_sent
  EditBox 75, 235, 105, 15, worker_signature
  ButtonGroup ButtonPressed
    CancelButton 36, 256, 70, 15
    OkButton 110, 256, 70, 15
  Text 10, 5, 175, 10, "Client did not respond to page for sameday interview"
  Text 33, 22, 44, 10, "Case Number"
  GroupBox 0, 40, 180, 70, "Client was Paged in the Lobby"
  Text 18, 59, 50, 10, "Interview Date:"
  Text 28, 75, 39, 8, "1st Page at:"
  Text 26, 92, 45, 15, "2nd Page at:"
  GroupBox 2, 117, 178, 79, "Phone Call to Client"
  Text 37, 147, 33, 10, "Called at:"
  Text 19, 164, 50, 12, "Phone Number"
  Text 13, 237, 60, 10, "Worker Signature"
EndDialog

'Connects to BlueZone default screen
EMConnect ""
EMFocus														

'Pulls case number from MAXIS if worker has already selected a case														
Call MAXIS_case_number_finder(case_number)

'Defaults the Interview Date to today's date
interview_date = date & ""

'Defaults the Client Phone number to the first phone number listed on MAXIS in STAT/ADDR
Call navigate_to_MAXIS_screen ("STAT", "ADDR")
EMReadScreen phone_01, 3, 17, 45
EMReadScreen phone_02, 3, 17, 51
EMReadScreen phone_03, 4, 17, 55
phone_number = phone_01 & "-" & phone_02 & "-" & phone_03 & ""

'Display's the Dialog Box to imput variable information - includes safeguards for mandatory fields
Do					
	Do
		Dialog no_show_dialog
		cancel_confirmation
		IF case_number = "" THEN MsgBox "You did not enter a case number. Please try again."
		IF interview_date = "" THEN MsgBox "You did not enter an Interview Date. Please try again."
		IF IsDate (interview_date) = False THEN MsgBox "Interview Date must be a date, please reenter."
		IF first_page = "" THEN MsgBox "Please enter the time of the 1st page in the lobby."
		IF second_page = "" THEN MsgBox "Please enter the time of the second page in the lobby - you must page your client at least twice"
		IF worker_signature = "" THEN MsgBox "You did not sign your case note. Please try again."
	Loop until case_number <> "" and interview_date <> "" and IsDate(interview_date) = True and first_page <> "" and second_page <> "" and worker_signature <> ""
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
Loop until abs(DateDiff("n", first_page, second_page))>=15 'and case_number <> "" and interview_date <> "" and IsDate(interview_date) = True and first_page <> "" and second_page <> "" and worker_signature <> ""

call check_for_MAXIS(False)	

'Pulls the application date listed on CASE/CURR from the CAF2 Pending line
Call navigate_to_MAXIS_screen("case", "curr")
EMReadScreen application_date, 8, 8, 29	

'Checks if worker wants script to send NOMI
IF nomi_sent = 1 THEN
	'Navigates to SPEC/LETR
	Call navigate_to_screen("SPEC", "LETR")

	'Checks to make sure we're past the SELF menu
	EMReadScreen still_self, 27, 2, 28
	If still_self = "Select Function Menu (SELF)" THEN script_end_proceedure("Unable to get past the SELF screen. Is your case in background?")									

	'Opens up the NOMI LETR. If it's unable the script will stop.
	EMWriteScreen "x", 7, 12
	transmit
	EMReadScreen LETR_check, 4, 2, 49
	IF LETR_check = "LETR" THEN script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")

	'Wries the info into the NOMI.
	EMWriteScreen "x", 7, 17
	call create_MAXIS_friendly_date(application_date, 0, 12, 38)
	call create_MAXIS_friendly_date(interview_date, 0, 14, 38)
	transmit
	PF4
End If

'Formats the page times and time called to standard hh:mm for case note
first_page = FormatDateTime (first_page, 4)
second_page = FormatDateTime (second_page ,4)
IF IsNumeric(time_called) = TRUE THEN
	time_called = FormatNumber (time_called, 2)
	time_called = FormatDateTime (time_called, 4)
End If

'Starts a Case Note
Call start_a_blank_case_note

'Writes the case note
call write_variable_in_CASE_NOTE("***Attempted to Page Client in Lobby for Interview - No Show***")
call write_bullet_and_variable_in_CASE_NOTE("Date of application", application_date)						
call write_bullet_and_variable_in_CASE_NOTE("Client was scheduled for interview", interview_date)				
call write_bullet_and_variable_in_CASE_NOTE("Paged client in the lobby to complete interview at", first_page & " & " & second_page)
IF pc_attempted = 1 THEN call write_bullet_and_variable_in_CASE_NOTE("Attempted to call client, no answer, called at provided number", phone_number & " at " & time_called)
IF left_vm = 1 THEN call write_variable_in_CASE_NOTE("* Left Voicemail for Client.")
IF nomi_sent = 1 THEN call write_variable_in_CASE_NOTE("* Sent NOMI to clt through SPEC/LETR.")			
IF potential_xfs = 1 THEN call write_variable_in_CASE_NOTE("* Case is Potentially XFS")
call write_variable_in_CASE_NOTE("---")			
call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure ("")

