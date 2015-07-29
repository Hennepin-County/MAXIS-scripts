'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - EVF RECEIVED.vbs"
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


BeginDialog EVF_received, 0, 0, 276, 170, "Employment Verification Form Received"
  EditBox 140, 5, 60, 15, date_received
  DropListBox 140, 25, 40, 15, "select"+chr(9)+"yes"+chr(9)+"no", signed_by_client
  DropListBox 140, 45, 40, 15, "select"+chr(9)+"yes"+chr(9)+"no", complete
  DropListBox 85, 65, 40, 15, "select"+chr(9)+"yes"+chr(9)+"no", faxed
  EditBox 180, 65, 60, 15, date_faxed
  EditBox 85, 85, 100, 15, fax_number
  EditBox 85, 105, 180, 15, notes
  EditBox 85, 125, 95, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 30, 150, 50, 15
    CancelButton 195, 150, 50, 15
  Text 10, 110, 70, 10, "Action taken / Notes:"
  Text 50, 50, 85, 10, "Completed by employer:"
  Text 15, 130, 60, 10, "Worker Signature:"
  Text 65, 10, 70, 10, "Date EVF received:"
  Text 70, 30, 55, 10, "Signed by client:"
  Text 135, 70, 40, 10, "Date faxed:"
  Text 20, 90, 65, 10, "Number faxed to:"
  Text 15, 70, 65, 10, "Faxed to employer:"
EndDialog


'connects to BlueZone and brings it forward
EMConnect ""
EMFocus

'checks that the worker is in MAXIS - allows them to get in MAXIS without ending the script
Call check_for_MAXIS(false)

'grabs the case number that is being worked on
call MAXIS_case_number_finder(case_number)

'starts the EVF received case note dialog
DO
	DO
		DO
			DO
				DO
					DO
						DO
							DO 
								'starts the EVF dialog
								Dialog EVF_received
								'asks if you want to cancel and if "yes" is selected sends StopScript
								cancel_confirmation 
								'checks that there is a date in the date received box
								IF IsDate (date_received) = FALSE THEN MsgBox "You must enter a date in mm/dd/yy for date received."
							LOOP UNTIL IsDate (date_received) = TRUE
							'checks if signed by client was selected
							IF Signed_by_client = "select" THEN MsgBox "You must select if signed by the client."
						LOOP UNTIL Signed_by_client <> "select"
						'checks if completed by employer was selected
						IF complete = "select" THEN MsgBox "You must select if completed by the employer."
					LOOP UNTIL complete <> "select"
					'checks if faxed to employer was selected
					IF faxed = "select" THEN MsgBox "You must select if faxed to the employer."
				LOOP UNTIL faxed <> "select"
				'checks that there is a faxed date entered if the EVF was faxed
				IF faxed = "yes" and IsDate (Date_faxed) = FALSE THEN MsgBox "You must enter a date in mm/dd/yy for date faxed."
			LOOP UNTIL faxed = "yes" and IsDate (Date_faxed) = TRUE or faxed = "no"
			'checks that there is a faxed number entered if the EVF was faxed
			IF faxed = "yes" and fax_number = "" THEN MsgBox "You must enter a fax number."
		LOOP UNTIL faxed = "yes" and fax_number <> "" or faxed = "no"
		'checks that notes were entered				
		IF notes = "" THEN MsgBox "You must enter action taken/notes."
	LOOP UNTIL notes <> ""
	'checks that the case note was signed
	IF worker_signature = "" THEN MsgBox "You must sign your case note!" 
LOOP UNTIL worker_signature <> "" 


'assigns a value to the EVF_status variable based on the value of the complete variable
IF complete = "yes" THEN EVF_status = "COMPLETE"
IF complete = "no" THEN EVF_status = "INCOMPLETE"


'checks that the worker is in MAXIS - allows them to get in MAXIS without ending the script
call check_for_MAXIS (false)

'starts a blank case note
call start_a_blank_case_note

'this enters the actual case note info 
call write_variable_in_CASE_NOTE("***EVF received " & date_received & " is " & EVF_status & "***")
call write_bullet_and_variable_in_CASE_NOTE("Date Received", date_received)
call write_bullet_and_variable_in_CASE_NOTE("Signed by client", signed_by_client)
call write_bullet_and_variable_in_CASE_NOTE("Completed by employer", complete)
call write_bullet_and_variable_in_CASE_NOTE("Faxed", faxed)
call write_bullet_and_variable_in_CASE_NOTE("Date Faxed", date_faxed)
call write_bullet_and_variable_in_CASE_NOTE("Fax number used", fax_number)
call write_bullet_and_variable_in_CASE_NOTE("Action Taken/Notes", notes)
call write_variable_in_CASE_NOTE ("---")
call write_variable_in_CASE_NOTE(worker_signature)





