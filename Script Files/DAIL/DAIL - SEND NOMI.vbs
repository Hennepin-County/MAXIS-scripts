'OPTION EXPLICIT
name_of_script = "DAIL - SEND NOMI.vbs"
start_time = timer

DIM name_of_script, start_time, FuncLib_URL, run_locally, default_directory, beta_agency, req, fso, row

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" OR default_directory = "" THEN
			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
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
'DIM ButtonPressed
'DIM interview_date
'DIM interview_time
'DIM recert_forms_confirm
'DIM result_of_msgbox

'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.

EMConnect ""

'Reading date and time of recert appt. from TIKL
'Dail message that should be read is: "~*~*~CLIENT HAD RECERT INTERVIEW AT" This is the part that is static in the DAIL message
EMReadScreen interview_date, 10, 17, 57
EMReadScreen interview_time, 8, 17, 71

'navigates to CASE/NOTE to user can see if interview has been completed or not
EMSendKey "n" 
transmit

'Msgbox asking the user to confirm if the client has sent a CAF or if no contact has been made by the client
recert_forms_confirm = MsgBox("The SNAP NOMI recertification SPEC/MEMO is ONLY to be sent when the SNAP recipient does not contact the agency about their recertification, and no CAF is received.  Press Yes if forms provided, OR contact was made by the recipient.") & _ 
	VbNewLine & ("Press No if no forms provided") & vbNewLine & ("Cancel to end the script.", vbYesNoCancel)
	If recert_forms_confirm = vbCancel then stopscript
	If recert_forms_confirm = vbYes then result_of_msgbox = TRUE
	If recert_forms_confirm = vbNo then result_of_msgbox = FALSE

If result_of_msgbox = TRUE then
	'Navigates into SPEC/MEMO
	call navigate_to_screen("SPEC", "MEMO")
	'Creates a new MEMO. If it's unable the script will stop.
	PF5
	'Writes the info into the MEMO.
	Call write_variable_in_SPEC_MEMO("************************************************************")
	Call write_variable_in_SPEC_MEMO("You have missed your SNAP interview that was scheduled for " & interview_date & " at " & interview_time & ".")
	Call write_variable_in_SPEC_MEMO("Please contact your worker at the telephone number listed below to reschedule the required SNAP interview.")
	Call write_variable_in_SPEC_MEMO("The renewal form, the interview by phone or in the office, and the mandatory verifications are needed to process your renewal must be completed by " & last_day_for_recert & " or your Food Support case will Auto-Close on this date.")
	Call write_variable_in_SPEC_MEMO("************************************************************")
	PF4	
	PF3

	'Navigates to a blank case note
	start_a_blank_case_note
	'Writes the case note
	Call write_variable_in_CASE_NOTE ("**Client missed SNAP recertification interview**")
	Call write_bullet_and_variable_in_CASE_NOTE "Appointment was scheduled for" & interview_date & " at " & interview_time & "."
	Call write_variable_in_CASE_NOTE "* A SNAP NOMI for recertifications SPEC/MEMO has been sent to the client informing them of their missed interview."
	Call write_variable_in_CASE_NOTE ("---")
	Call write_variable_in_CASE_NOTE (worker_signature)
	PF3
	script_end_procedure("Success! A SNAP NOMI for recertifications SPEC/MEMO has been sent.")
END IF 

If result_of_msgbox = FALSE
	'Navigates to a blank case note
	start_a_blank_case_note
	'Writes the case note
	Call write_variable_in_CASE_NOTE ("**Client missed SNAP recertification interview**")
	Call write_bullet_and_variable_in_CASE_NOTE "Appointment was scheduled for" & interview_date & " at " & interview_time & "."
	Call write_variable_in_CASE_NOTE ("* A SNAP NOMI for recertifications SPEC/MEMO HAS NOT been sent. Per POLI/TEMP TE02.05.15: When there is no request for further assistance the client will receive the proper closing.")
	Call write_variable_in_CASE_NOTE ("---")
	Call write_variable_in_CASE_NOTE (worker_signature)
	PF3
	script_end_procedure("Success! A SNAP NOMI for recertifications SPEC/MEMO has NOT been sent.")
END IF 