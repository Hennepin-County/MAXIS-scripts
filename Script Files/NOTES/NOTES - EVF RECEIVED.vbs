'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - EVF RECEIVED.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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

'Required for statistical purposes==========================================================================================
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 120          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

BeginDialog case_number_dlg, 0, 0, 206, 75, "Enter a Case Number"
  EditBox 70, 10, 70, 15, case_number
  EditBox 65, 30, 30, 15, benefit_month
  EditBox 155, 30, 30, 15, benefit_year
  ButtonGroup ButtonPressed
    OkButton 100, 55, 50, 15
    CancelButton 150, 55, 50, 15
  Text 10, 15, 50, 10, "Case Number"
  Text 10, 35, 50, 10, "Benefit Month"
  Text 105, 35, 45, 10, "Benefit Year"
EndDialog

BeginDialog EVF_received, 0, 0, 276, 200, "Employment Verification Form Received"
  EditBox 140, 4, 60, 16, date_received
  EditBox 54, 28, 72, 16, client
  EditBox 190, 28, 70, 16, employer
  DropListBox 144, 52, 50, 16, "Select one"+chr(9)+"yes"+chr(9)+"no", signed_by_client
  DropListBox 144, 70, 50, 16, "Select one"+chr(9)+"yes"+chr(9)+"no", complete
  DropListBox 92, 92, 48, 16, "Select one"+chr(9)+"yes"+chr(9)+"no", info
  EditBox 206, 88, 60, 16, info_date
  EditBox 82, 110, 100, 16, request_info
  EditBox 82, 132, 180, 16, notes
  EditBox 82, 154, 96, 16, worker_signature
  ButtonGroup ButtonPressed
    OkButton 28, 176, 50, 16
    CancelButton 200, 176, 50, 16
  Text 12, 158, 60, 10, "Worker Signature:"
  Text 72, 10, 62, 10, "Date EVF received:"
  Text 82, 56, 56, 10, "Signed by client:"
  Text 146, 94, 56, 10, "Date Requested:"
  Text 14, 114, 64, 10, "Info Requested via:"
  Text 4, 94, 86, 10, "Additional Info Requested"
  Text 18, 34, 30, 14, "MEMB #:"
  Text 136, 32, 52, 12, "Employer name:"
  Text 56, 72, 80, 10, "Completed by employer:"
  Text 10, 138, 68, 10, "Action taken / Notes:"
EndDialog

'connects to BlueZone and brings it forward
EMConnect ""
EMFocus

'checks that the worker is in MAXIS - allows them to get in MAXIS without ending the script
Call check_for_MAXIS(false)

'grabs the case number and benefit month/year that is being worked on
call MAXIS_case_number_finder(case_number)
EMReadScreen at_self, 4, 2, 50
IF at_self = "SELF" THEN 
	EMReadScreen benefit_month, 2, 20, 43
	IF len(benefit_month) <> 2 THEN benefit_month = "0" & benefit_month
	EMReadScreen benefit_year, 2, 20, 46
ELSE
	CALL find_variable("Month: ", benefit_month, 2)
	IF benefit_month <> "  " THEN CALL find_variable("Month: " & benefit_month & " ", benefit_year, 2)
END IF

' >>>>> GATHERING & CONFIRMING THE MAXIS CASE NUMBER <<<<<
DO
	err_msg = ""
	DIALOG case_number_dlg
		IF ButtonPressed = 0 THEN stopscript
		IF case_number = "" OR (case_number <> "" AND len(case_number) > 8) OR (case_number <> "" AND IsNumeric(case_number) = False) THEN err_msg = err_msg & vbCr & "* Please enter a valid case number."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."		
LOOP UNTIL err_msg = ""

CALL check_for_MAXIS(False)

'starts the EVF received case note dialog
DO
	err_msg = ""
	'starts the EVF dialog
	Dialog EVF_received
	'asks if you want to cancel and if "yes" is selected sends StopScript
	cancel_confirmation 
	'checks that there is a date in the date received box
	IF IsDate (date_received) = FALSE THEN err_msg = err_msg & vbCr & "You must enter a date in mm/dd/yy for date received."
	'checks if the client name has been entered
	IF client = "" THEN err_msg = err_msg & vbCr & "You must enter the MEMB #."
	'checks if the employer name has been entered
	IF employer = "" THEN err_msg = err_msg & vbCr & "You must enter the employers name."
	'checks if signed by client was selected
	IF Signed_by_client = "Select one" THEN err_msg = err_msg & vbCr & "You must select if signed by the client."
	'checks if completed by employer was selected
	IF complete = "Select one" THEN err_msg = err_msg & vbCr & "You must select if completed by the employer."
	'checks if additional info was requested 
	IF info = "Select one" THEN err_msg = err_msg & vbCr & "You must select if additional info was requested."
	'checks that there is a info request date entered if the it was requested
	IF info = "yes" and IsDate (info_date) = FALSE THEN err_msg = err_msg & vbCr & "You must enter a date in mm/dd/yy that additional info was requested."
	'checks that there is a method of inquiry entered if additional info was requested
	IF info = "yes" and request_info = "" THEN err_msg = err_msg & vbCr & "You must enter the method used to request additional info."
	'checks that notes were entered				
	IF notes = "" THEN err_msg = err_msg & vbCr & "You must enter action taken/notes."
	'checks that the case note was signed
	IF worker_signature = "" THEN err_msg = err_msg & vbCr & "You must sign your case note!" 
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""

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
call write_variable_in_CASE_NOTE("* MEMB #/Employer: MEMB " & client & " at " & employer)
call write_bullet_and_variable_in_CASE_NOTE("Signed by client", signed_by_client)
call write_bullet_and_variable_in_CASE_NOTE("Completed by employer", complete)
	'case note changes based on if additional info was requested
	IF info = "yes" then call write_variable_in_CASE_NOTE ("* Additional Info requested: " & info & " on " & info_date)
	IF info = "no" then call write_variable_in_CASE_NOTE ("* Additional Info requested: " & info)
call write_bullet_and_variable_in_CASE_NOTE("Request method used", request_info)
call write_bullet_and_variable_in_CASE_NOTE("Action Taken/Notes", notes)
	'case notes that a TIKL was set if additional information was requested
	IF info = "yes" THEN call write_variable_in_CASE_NOTE ("***TIKLed for 10 day return.***")
call write_variable_in_CASE_NOTE ("---")
call write_variable_in_CASE_NOTE(worker_signature)

'Checks if additional info is yes and sets a TIKL for the return of the info
IF info = "yes" THEN 
	call navigate_to_MAXIS_screen("dail", "writ")

	'The following will generate a TIKL formatted date for 10 days from now.
	call create_MAXIS_friendly_date(date, 10, 5, 18)

	'Writing in the rest of the TIKL.
	call write_variable_in_TIKL("Additional info requested after an EVF being rec'd should have returned by now. If not received, take appropriate action. (TIKL auto-generated from script)." )
	transmit
	PF3

	'Success message
	MsgBox "Success! TIKL has been sent for 10 days from now for the additional information requested."

End if
script_end_procedure("")
