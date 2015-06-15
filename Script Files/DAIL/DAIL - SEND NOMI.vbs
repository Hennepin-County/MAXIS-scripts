'OPTION EXPLICIT

name_of_script = "DAIL - SEND NOMI.vbs"
start_time = timer

DIM name_of_script
DIM start_time
DIM FuncLib_URL
DIM run_locally
DIM default_directory
DIM beta_agency
DIM req
DIM fso
DIM row

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
'END OF GLOBAL VARIABLES----------------------------------------------------------------------------------------------------

'Declaring variables
DIM ButtonPressed
DIM missed_interview_dialog
DIM completed_interview_dialog

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog missed_interview_dialog, 0, 0, 226, 45, "Missed interview dialog"
  ButtonGroup ButtonPressed
    OkButton 55, 25, 50, 15
    CancelButton 110, 25, 50, 15
  Text 5, 10, 225, 10, "It appears this case has missed an interview.  Press OK to confirm."
EndDialog

BeginDialog completed_interview_dialog, 0, 0, 241, 45, "Completed interview dialog"
  ButtonGroup ButtonPressed
    OkButton 55, 25, 50, 15
    CancelButton 110, 25, 50, 15
  Text 5, 10, 230, 10, "It appears this case has completed an interview.  Press OK to confirm."
EndDialog

'DAIL SCRUBBER START----------------------------------------------------------------------------------------------------
EMConnect ""

'Reading date and time of recert appt. from TIKL
EMReadScreen date_of_missed_interview   *********************************************************
EMReadScreen time_of_missed_interview   *********************************************************
PF3 

'navigates to CASE/NOTE
EMSendKey "n" 
transmit

'checking to see if case note exists prior to recertification date listed in TIKL
EMReadScreen post_scheduled_appt_note_check, 8, 5, 6


If post_scheduled_appt_note_check = appt_date THEN
	completed_interview_dialog
	cancel_confirmation
End IF

If post_scheduled_appt_note_check <> appt_date or date after**** THEN 
	DO
		missed_interview_dialog
		cancel_confimation
		Call proceed_confirmation(go_to_case_note)
	Loop until go_to_case_note = TRUE
	
	'THE CASE NOTE----------------------------------------------------------------------------------------------------
	'Navigates to a blank case note
	Call start_a_blank_case_note
	'Writes the case note
	Call write_variable_in_CASE_NOTE ("**Client missed SNAP recertification interview**")
	Call write_bullet_and_variable_in_CASE_NOTE "Appointment was scheduled for" & date_of_missed_interview & " at " & time_of_missed_interview & "."
	Call write_variable_in_CASE_NOTE "* A SPEC/MEMO has been sent to the client informing them of missed interview."
	Call write_variable_in_CASE_NOTE ("---")
	Call write_variable_in_CASE_NOTE (worker_signature)
	
	
	'THE ER NOMI SPEC/MEMO----------------------------------------------------------------------------------------------------
	'Navigates into SPEC/MEMO
	call navigate_to_screen("SPEC", "MEMO")
	'Creates a new MEMO. If it's unable the script will stop.
	PF5
	EMReadScreen memo_display_check, 12, 2, 33
	If memo_display_check = "Memo Display" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")
	EMWriteScreen "x", 5, 10
	transmit
	'Writes the info into the MEMO.
	Call write_variable_in_SPEC_MEMO("************************************************************")
	Call write_variable_in_SPEC_MEMO("You have missed your Food Support interview that was scheduled for " & date_of_missed_interview & " at " & time_of_missed_interview & ".")
	Call write_variable_in_SPEC_MEMO("Please contact your worker at the telephone number listed below to reschedule the required Food Support interview.")
	Call write_variable_in_SPEC_MEMO("The Combined Application Form (DHS-5223), the interview by phone or in the office, and the mandatory verifications needed to process your recertification must be completed by " & last_day_for_recert & " or your Food Support case will Auto-Close on this date." & "<newline>"
	Call write_variable_in_SPEC_MEMO("************************************************************")
	PF4	
	PF3
	'Pop up for worker informing them NOMI has been sent for missed SNAP/MFIP ER	
	MsgBox "Success! A SPEC/MEMO has been sent with the correct language for a missed SNAP recert. A case note has been made."
END IF

script_end_procedure("")

